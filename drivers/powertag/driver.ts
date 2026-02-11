'use strict';

import Homey from 'homey';
import * as net from 'net';
import * as Modbus from 'jsmodbus';
import {
  POWERTAG_MODELS,
  getCapabilitiesForModel,
  getCapabilityOptionsForModel,
} from '../../lib/PowerTagRegistry';
import { readDeviceType, readDeviceName } from '../../lib/ModbusHelpers';
import type { PowerTagSettings, PowerTagStore, PowerTagDeviceData } from '../../lib/types';

/**
 * Base slave ID for wireless devices on a Smartlink SI D / PowerTag Link.
 * Device slot 0 → slave 150, slot 1 → slave 151, etc.
 */
const SMARTLINK_SLAVE_BASE = 150;

class PowerTagDriver extends Homey.Driver {

  async onInit(): Promise<void> {
    this.log('PowerTag driver initialized');
  }

  async onPair(session: any): Promise<void> {
    let gatewaySettings: PowerTagSettings | null = null;

    session.setHandler('gateway_settings', async (data: PowerTagSettings) => {
      this.log(`Validating gateway at ${data.address}:${data.port}`);
      gatewaySettings = data;

      // Validate connection to the gateway
      await this.validateGateway(data.address, data.port);
      this.log('Gateway validation successful');

      // Navigate to list_devices from the driver side
      await session.nextView();
      return true;
    });

    session.setHandler('list_devices', async () => {
      if (!gatewaySettings) {
        throw new Error('No gateway configured');
      }
      return this.discoverDevices(gatewaySettings);
    });
  }

  /**
   * Validate that we can connect to the gateway via TCP.
   */
  private async validateGateway(host: string, port: number): Promise<void> {
    return new Promise((resolve, reject) => {
      const socket = new net.Socket();
      const timeout = setTimeout(() => {
        socket.destroy();
        reject(new Error(`Connection timeout to ${host}:${port}`));
      }, 5000);

      socket.connect({ host, port }, () => {
        clearTimeout(timeout);
        socket.destroy();
        resolve();
      });

      socket.on('error', (err: Error) => {
        clearTimeout(timeout);
        socket.destroy();
        reject(new Error(`Cannot connect to ${host}:${port}: ${err.message}`));
      });
    });
  }

  /** Set the unit ID on the client for the next request(s). */
  private setUnitId(client: InstanceType<typeof Modbus.client.TCP>, unitId: number): void {
    (client as any)._unitId = unitId;
    (client as any)._requestHandler._unitId = unitId;
  }

  /**
   * Discover PowerTag devices on the gateway.
   *
   * Scans the Smartlink range (150-169) first since that's where most
   * gateways place devices, then extends to 100-149 if nothing is found
   * (for Panel Servers that use the full 100-199 range).
   */
  private async discoverDevices(settings: PowerTagSettings): Promise<any[]> {
    const socket = new net.Socket();

    // Create the Modbus client BEFORE connecting — jsmodbus must see the
    // socket 'connect' event to transition its internal state to 'online'.
    // 500ms timeout: real devices respond in ~5-50ms, the gateway returns
    // Modbus exceptions in <10ms for non-existent registers.
    const client = new Modbus.client.TCP(socket, SMARTLINK_SLAVE_BASE, 500);

    try {
      await new Promise<void>((resolve, reject) => {
        const timeout = setTimeout(() => {
          socket.destroy();
          reject(new Error('Discovery connection timeout'));
        }, 10000);

        socket.connect({ host: settings.address, port: settings.port }, () => {
          clearTimeout(timeout);
          resolve();
        });

        socket.on('error', (err: Error) => {
          clearTimeout(timeout);
          reject(err);
        });
      });

      // Scan Smartlink range first (150-169), then Panel Server range (100-149)
      let devices = await this.scanSlaveRange(client, settings, 150, 170);
      if (devices.length === 0) {
        this.log('No devices at 150-169, scanning 100-149...');
        devices = await this.scanSlaveRange(client, settings, 100, 150);
      }

      this.log(`Discovery complete: found ${devices.length} devices`);
      return devices;
    } finally {
      socket.destroy();
    }
  }

  /**
   * Scan a range of slave IDs for devices by reading register 31024 (device type).
   * Stops early after 10 consecutive timeouts once at least one device is found.
   */
  private async scanSlaveRange(
    client: InstanceType<typeof Modbus.client.TCP>,
    settings: PowerTagSettings,
    startId: number,
    endId: number,
  ): Promise<any[]> {
    const devices: any[] = [];

    this.log(`Scanning slave IDs ${startId}-${endId - 1}...`);
    let consecutiveTimeouts = 0;
    for (let slaveId = startId; slaveId < endId; slaveId++) {
      try {
        this.setUnitId(client, slaveId);
        const typeId = await readDeviceType(client);
        consecutiveTimeouts = 0;

        if (typeId === 0 || typeId === 65535) continue;

        this.log(`Slave ${slaveId}: typeId=${typeId}`);

        const modelConfig = POWERTAG_MODELS.get(typeId);
        if (!modelConfig) {
          this.log(`Unknown device type ${typeId} at slave ${slaveId}, skipping`);
          continue;
        }

        // Try to read the user-configured name (may not be supported on all gateways)
        let deviceName = '';
        try {
          deviceName = await readDeviceName(client);
        } catch {
          // Name register not supported — use model name fallback
        }

        devices.push({
          name: deviceName || `${modelConfig.name} (${slaveId})`,
          data: { id: `${settings.address}:${settings.port}:${slaveId}` } as PowerTagDeviceData,
          settings: {
            address: settings.address,
            port: settings.port,
            polling: settings.polling,
          },
          store: { slaveId, typeId, model: modelConfig.model } as PowerTagStore,
          capabilities: getCapabilitiesForModel(modelConfig),
          capabilitiesOptions: getCapabilityOptionsForModel(modelConfig),
        });

        this.log(`Found ${modelConfig.model} at slave ${slaveId}`);
      } catch {
        consecutiveTimeouts++;
        if (consecutiveTimeouts >= 10 && devices.length > 0) {
          this.log(`Stopping scan early after 10 consecutive timeouts at slave ${slaveId}`);
          break;
        }
      }
    }

    return devices;
  }

}

module.exports = PowerTagDriver;
