'use strict';

import * as net from 'net';
import * as Modbus from 'jsmodbus';

/** Minimal interface for the Homey instance (avoids namespace-as-type issue) */
interface HomeyInstance {
  setTimeout(callback: Function, ms: number, ...args: any[]): NodeJS.Timeout;
  clearTimeout(timeoutId: any): void;
  setInterval(callback: Function, ms: number, ...args: any[]): NodeJS.Timeout;
  clearInterval(timeoutId: any): void;
  app: any;
}

interface QueuedOperation {
  unitId: number;
  execute: (client: InstanceType<typeof Modbus.client.TCP>) => Promise<any>;
  resolve: (value: any) => void;
  reject: (reason: any) => void;
}

interface ConnectionEntry {
  socket: net.Socket;
  client: InstanceType<typeof Modbus.client.TCP>;
  refCount: number;
  connected: boolean;
  connecting: Promise<void> | null;
  queue: QueuedOperation[];
  processing: boolean;
  host: string;
  port: number;
}

const CONNECT_TIMEOUT = 10_000;
const RECONNECT_DELAY = 30_000;
const OPERATION_TIMEOUT = 5_000;

/**
 * Manages a pool of Modbus TCP connections, one per gateway IP:port.
 * All PowerTag devices on the same gateway share a single TCP socket.
 * Operations are queued and executed sequentially (Modbus is request-response).
 *
 * Important: the jsmodbus TCP client must be created BEFORE socket.connect()
 * so it can observe the 'connect' event. We create one client per connection
 * and swap the unit ID before each operation.
 */
export class ModbusConnectionManager {

  private connections: Map<string, ConnectionEntry> = new Map();
  private homey: HomeyInstance;

  constructor(homey: HomeyInstance) {
    this.homey = homey;
  }

  private getKey(host: string, port: number): string {
    return `${host}:${port}`;
  }

  /**
   * Create a new socket + Modbus client pair.
   * The client is created before connect so it sees the 'connect' event.
   */
  private createSocketAndClient(host: string, port: number): { socket: net.Socket; client: InstanceType<typeof Modbus.client.TCP> } {
    const socket = new net.Socket();
    // Create client before connecting — jsmodbus needs to see the 'connect' event
    const client = new Modbus.client.TCP(socket, 1);
    return { socket, client };
  }

  /**
   * Acquire a reference to a connection. Creates the socket if needed.
   * Call release() when the device is deleted/uninitialized.
   */
  async acquire(host: string, port: number): Promise<void> {
    const key = this.getKey(host, port);
    let entry = this.connections.get(key);

    if (!entry) {
      const { socket, client } = this.createSocketAndClient(host, port);
      entry = {
        socket,
        client,
        refCount: 0,
        connected: false,
        connecting: null,
        queue: [],
        processing: false,
        host,
        port,
      };
      this.setupSocketHandlers(key, entry);
      this.connections.set(key, entry);
    }

    entry.refCount++;

    if (!entry.connected && !entry.connecting) {
      await this.connect(entry);
    } else if (entry.connecting) {
      await entry.connecting;
    }
  }

  /**
   * Release a reference. Destroys the socket when refCount reaches 0.
   */
  release(host: string, port: number): void {
    const key = this.getKey(host, port);
    const entry = this.connections.get(key);
    if (!entry) return;

    entry.refCount--;
    if (entry.refCount <= 0) {
      entry.socket.destroy();
      this.connections.delete(key);
    }
  }

  /**
   * Execute a Modbus operation on a specific unit ID.
   * Returns a promise that resolves with the operation result.
   */
  async execute<T>(
    host: string,
    port: number,
    unitId: number,
    operation: (client: InstanceType<typeof Modbus.client.TCP>) => Promise<T>,
  ): Promise<T> {
    const key = this.getKey(host, port);
    const entry = this.connections.get(key);

    if (!entry) {
      throw new Error(`No connection for ${key}`);
    }

    // If not connected, attempt reconnect
    if (!entry.connected) {
      if (!entry.connecting) {
        entry.connecting = this.connect(entry).catch(() => {});
      }
      await entry.connecting;
      if (!entry.connected) {
        throw new Error(`Not connected to ${key}`);
      }
    }

    return new Promise<T>((resolve, reject) => {
      entry.queue.push({
        unitId,
        execute: operation as any,
        resolve,
        reject,
      });

      if (!entry.processing) {
        this.processQueue(entry);
      }
    });
  }

  private setupSocketHandlers(key: string, entry: ConnectionEntry): void {
    entry.socket.on('close', () => {
      entry.connected = false;
      entry.connecting = null;

      // Reject pending operations
      for (const op of entry.queue) {
        op.reject(new Error('Connection closed'));
      }
      entry.queue = [];
      entry.processing = false;

      // Schedule reconnect if still referenced
      if (entry.refCount > 0) {
        this.homey.setTimeout(() => {
          if (entry.refCount > 0 && !entry.connected && !entry.connecting) {
            this.connect(entry).catch(() => {});
          }
        }, RECONNECT_DELAY);
      }
    });

    entry.socket.on('error', (err: Error) => {
      (this.homey.app as any)?.log?.(`Modbus socket error [${key}]: ${err.message}`);
    });
  }

  private async connect(entry: ConnectionEntry): Promise<void> {
    if (entry.connecting) return entry.connecting;

    entry.connecting = new Promise<void>((resolve, reject) => {
      // Create a fresh socket + client if the old socket is destroyed
      if (entry.socket.destroyed) {
        const { socket, client } = this.createSocketAndClient(entry.host, entry.port);
        entry.socket = socket;
        entry.client = client;
        this.setupSocketHandlers(this.getKey(entry.host, entry.port), entry);
      }

      const timeout = setTimeout(() => {
        entry.socket.destroy();
        entry.connecting = null;
        reject(new Error(`Connection timeout to ${entry.host}:${entry.port}`));
      }, CONNECT_TIMEOUT);

      entry.socket.connect({ host: entry.host, port: entry.port }, () => {
        clearTimeout(timeout);
        entry.connected = true;
        entry.connecting = null;
        resolve();
      });

      entry.socket.once('error', (err: Error) => {
        clearTimeout(timeout);
        entry.connecting = null;
        reject(err);
      });
    });

    return entry.connecting;
  }

  private async processQueue(entry: ConnectionEntry): Promise<void> {
    if (entry.processing) return;
    entry.processing = true;

    while (entry.queue.length > 0 && entry.connected) {
      const op = entry.queue.shift()!;

      try {
        // Swap the unit ID on the shared client for this operation
        (entry.client as any)._unitId = op.unitId;
        (entry.client as any)._requestHandler._unitId = op.unitId;

        const result = await Promise.race([
          op.execute(entry.client),
          new Promise<never>((_, reject) =>
            setTimeout(() => reject(new Error('Operation timeout')), OPERATION_TIMEOUT),
          ),
        ]);
        op.resolve(result);
      } catch (err) {
        op.reject(err);
      }
    }

    entry.processing = false;
  }

  /** Destroy all connections. Called on app uninit. */
  destroy(): void {
    for (const [, entry] of this.connections) {
      entry.socket.destroy();
      for (const op of entry.queue) {
        op.reject(new Error('Manager destroyed'));
      }
    }
    this.connections.clear();
  }
}
