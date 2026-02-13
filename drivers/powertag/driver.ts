'use strict';

import Homey from 'homey';
import * as net from 'net';
import * as Modbus from 'jsmodbus';
import {
  POWERTAG_MODELS,
  getModelByReference,
  getCapabilitiesForModel,
  getCapabilityOptionsForModel,
} from '../../lib/PowerTagRegistry';
import {
  readDeviceType,
  readDeviceName,
  readPanelServerDeviceAddresses,
  readCommercialReference,
} from '../../lib/ModbusHelpers';
import type { PowerTagSettings, PowerTagStore, PowerTagDeviceData, PowerTagModelConfig } from '../../lib/types';

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
   * Tries PAS600 Panel Server discovery first (reads device address table
   * from gateway unit 255). If that fails, falls back to Smartlink-style
   * scanning of unit ID ranges.
   */
  private async discoverDevices(settings: PowerTagSettings): Promise<any[]> {
    const socket = new net.Socket();

    // Create the Modbus client BEFORE connecting — jsmodbus must see the
    // socket 'connect' event to transition its internal state to 'online'.
    // 500ms timeout: real devices respond in ~5-50ms, the gateway returns
    // Modbus exceptions in <10ms for non-existent registers.
    const client = new Modbus.client.TCP(socket, 1, 500);

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

      // Try Panel Server (PAS600) discovery first — reads device addresses
      // from gateway unit 255. Fast when supported (~1-2s), single timeout
      // (~500ms) when not.
      let devices: any[] = [];
      try {
        devices = await this.discoverViaPanelServer(client, settings);
        this.log(`Panel Server discovery found ${devices.length} devices`);
      } catch (err) {
        this.log('Panel Server discovery not available, falling back to Smartlink scan');
      }

      // Fall back to Smartlink-style range scanning if PAS600 didn't find anything
      if (devices.length === 0) {
        const smartlink = await this.scanUnitRange(client, settings, 150, 170);
        const lower = await this.scanUnitRange(client, settings, 100, 150, 3);
        const upper = await this.scanUnitRange(client, settings, 170, 200, 3);
        devices = [...smartlink, ...lower, ...upper];
        this.log(`Smartlink scan found ${devices.length} devices`);
      }

      this.log(`Discovery complete: found ${devices.length} devices`);
      return devices;
    } finally {
      socket.destroy();
    }
  }

  /**
   * Discover devices via PAS600 Panel Server.
   *
   * Reads the device address table from gateway unit 255, then queries
   * each device for its commercial reference string to identify the model.
   */
  private async discoverViaPanelServer(
    client: InstanceType<typeof Modbus.client.TCP>,
    settings: PowerTagSettings,
  ): Promise<any[]> {
    // Read device address table from gateway (unit 255)
    this.setUnitId(client, 255);
    const addressMap = await readPanelServerDeviceAddresses(client);
    this.log(`Panel Server address table: ${addressMap.size} devices found`);

    const devices: any[] = [];

    for (const [slot, unitId] of addressMap) {
      try {
        this.setUnitId(client, unitId);

        // Read commercial reference to identify the model
        const reference = await readCommercialReference(client);
        if (!reference) {
          this.log(`Slot ${slot} (unit ${unitId}): empty reference, skipping`);
          continue;
        }

        const modelConfig = getModelByReference(reference);
        if (!modelConfig) {
          this.log(`Slot ${slot} (unit ${unitId}): unknown reference "${reference}", skipping`);
          continue;
        }

        // Try to read the user-configured name
        let deviceName = '';
        try {
          deviceName = await readDeviceName(client);
        } catch {
          // Name register not supported — use model name fallback
        }

        devices.push(this.buildDeviceEntry(
          settings, unitId, modelConfig.typeId, modelConfig, deviceName,
        ));

        this.log(`Found ${modelConfig.model} "${reference}" at unit ${unitId} (slot ${slot})`);
      } catch (err) {
        this.log(`Slot ${slot} (unit ${unitId}): read failed, skipping`);
      }
    }

    return devices;
  }

  /**
   * Scan a range of unit IDs for devices by reading register 31024 (device type).
   * Stops early after consecutive timeouts to keep Smartlink scans fast.
   */
  private async scanUnitRange(
    client: InstanceType<typeof Modbus.client.TCP>,
    settings: PowerTagSettings,
    startId: number,
    endId: number,
    maxConsecutiveTimeouts = 10,
  ): Promise<any[]> {
    const devices: any[] = [];

    this.log(`Scanning unit IDs ${startId}-${endId - 1}...`);
    let consecutiveTimeouts = 0;
    for (let unitId = startId; unitId < endId; unitId++) {
      try {
        this.setUnitId(client, unitId);
        const typeId = await readDeviceType(client);
        consecutiveTimeouts = 0;

        if (typeId === 0 || typeId === 65535) continue;

        this.log(`Unit ${unitId}: typeId=${typeId}`);

        const modelConfig = POWERTAG_MODELS.get(typeId);
        if (!modelConfig) {
          this.log(`Unknown device type ${typeId} at unit ${unitId}, skipping`);
          continue;
        }

        // Try to read the user-configured name (may not be supported on all gateways)
        let deviceName = '';
        try {
          deviceName = await readDeviceName(client);
        } catch {
          // Name register not supported — use model name fallback
        }

        devices.push(this.buildDeviceEntry(settings, unitId, typeId, modelConfig, deviceName));

        this.log(`Found ${modelConfig.model} at unit ${unitId}`);
      } catch {
        consecutiveTimeouts++;
        if (consecutiveTimeouts >= maxConsecutiveTimeouts) {
          this.log(`Stopping scan after ${maxConsecutiveTimeouts} consecutive timeouts at unit ${unitId}`);
          break;
        }
      }
    }

    return devices;
  }

  /** Build a Homey device entry for the pairing list. */
  private buildDeviceEntry(
    settings: PowerTagSettings,
    unitId: number,
    typeId: number,
    modelConfig: PowerTagModelConfig,
    deviceName: string,
  ): any {
    return {
      name: deviceName || `${modelConfig.name} (${unitId})`,
      data: { id: `${settings.address}:${settings.port}:${unitId}` } as PowerTagDeviceData,
      settings: {
        address: settings.address,
        port: settings.port,
        polling: settings.polling,
      },
      store: { unitId, typeId, model: modelConfig.model } as PowerTagStore,
      capabilities: getCapabilitiesForModel(modelConfig),
      capabilitiesOptions: getCapabilityOptionsForModel(modelConfig),
    };
  }

}

module.exports = PowerTagDriver;
