'use strict';

import * as Modbus from 'jsmodbus';
import { VoltageMode, EnergyPollResult, HeatTagPollResult, Control2DIPollResult, ControlIOPollResult } from './types';

// Register addresses (0-based, FC3 Read Holding Registers)

const REG_CURRENT_L1   = 2999;
const REG_VOLTAGE_LN_1 = 3027;
const REG_VOLTAGE_LL_1 = 3019;
const REG_POWER_L1     = 3053;
const REG_POWER_FACTOR = 3083;
const REG_FREQUENCY    = 3109;
const REG_TEMPERATURE  = 3131;
const REG_ENERGY_TOTAL = 3203;

/** Device type ID register, used during discovery (Smartlink only, hex 0x7930) */
export const REG_DEVICE_TYPE = 31024;

/** User-configured device name (ASCII, max 20 chars, 10 registers, hex 0x7918) */
const REG_DEVICE_NAME = 31000;

/** PAS600 Panel Server: device address table base at unit 255 (hex 0x01F8) */
const REG_PAS_DEVICE_ADDRESS = 0x01F8; // 504

/** PAS600 Panel Server: commercial reference string per device (hex 0x7954, 16 regs) */
const REG_COMMERCIAL_REF = 0x7954; // 31060

/**
 * Read all measurement registers for a PowerTag device.
 *
 * Groups adjacent registers into contiguous block reads to minimize
 * Modbus transactions:
 *   Block 1: Current L1/L2/L3 (2999-3004, 6 regs)
 *   Block 2: Voltage (3019-3024 L-L or 3027-3032 L-N, 6 regs)
 *   Block 3: Power L1/L2/L3 + Total (3053-3060, 8 regs)
 *   Block 4: Power Factor (3083-3084, 2 regs)
 *   Block 5: Frequency (3109-3110, 2 regs)
 *   Block 6: Temperature (3131-3132, 2 regs)
 *   Block 7: Total Energy (3203-3206, 4 regs)
 */
export async function readAllRegisters(
  client: InstanceType<typeof Modbus.client.TCP>,
  voltageMode: VoltageMode,
): Promise<EnergyPollResult> {
  const voltStart = voltageMode === 'L-N' ? REG_VOLTAGE_LN_1 : REG_VOLTAGE_LL_1;

  const [currResp, voltResp, powerResp, pfResp, freqResp, tempResp, energyResp] =
    await Promise.all([
      client.readHoldingRegisters(REG_CURRENT_L1, 6),
      client.readHoldingRegisters(voltStart, 6),
      client.readHoldingRegisters(REG_POWER_L1, 8),
      client.readHoldingRegisters(REG_POWER_FACTOR, 2),
      client.readHoldingRegisters(REG_FREQUENCY, 2),
      client.readHoldingRegisters(REG_TEMPERATURE, 2),
      client.readHoldingRegisters(REG_ENERGY_TOTAL, 4),
    ]);

  const currBuf   = currResp.response.body.valuesAsBuffer;
  const voltBuf   = voltResp.response.body.valuesAsBuffer;
  const powerBuf  = powerResp.response.body.valuesAsBuffer;
  const pfBuf     = pfResp.response.body.valuesAsBuffer;
  const freqBuf   = freqResp.response.body.valuesAsBuffer;
  const tempBuf   = tempResp.response.body.valuesAsBuffer;
  const energyBuf = energyResp.response.body.valuesAsBuffer;

  return {
    currentL1:   currBuf.readFloatBE(0),
    currentL2:   currBuf.readFloatBE(4),
    currentL3:   currBuf.readFloatBE(8),
    voltagePh1:  voltBuf.readFloatBE(0),
    voltagePh2:  voltBuf.readFloatBE(4),
    voltagePh3:  voltBuf.readFloatBE(8),
    powerL1:     powerBuf.readFloatBE(0),
    powerL2:     powerBuf.readFloatBE(4),
    powerL3:     powerBuf.readFloatBE(8),
    totalPower:  powerBuf.readFloatBE(12),
    powerFactor: pfBuf.readFloatBE(0),
    frequency:   freqBuf.readFloatBE(0),
    temperature: tempBuf.readFloatBE(0),
    totalEnergy: Number(energyBuf.readBigInt64BE(0)) / 1000, // Wh -> kWh
  };
}

// HeatTag registers
const REG_HEATTAG_TEMP     = 4001;  // Float32: temperature °C
const REG_HEATTAG_HUMIDITY = 4007;  // Float32: relative humidity (0.50 = 50%)
const REG_HEATTAG_ALARM    = 3323;  // UINT16: alarm level (0=None, 1=Low, 2=Medium, 3=High)

// Control module registers
const REG_DI1_STATUS = 34065;  // UINT16: DI1 input status (0=On, 1=Off)
const REG_DI2_STATUS = 34165;  // UINT16: DI2 input status (0=On, 1=Off)
const REG_DO1_CMD    = 37051;  // UINT16 R/W: command (0=None, 1=Off, 2=On)
const REG_DO1_STATUS = 37052;  // UINT16: output status (0=Off, 1=On)

/**
 * Read HeatTag sensor registers: temperature, humidity, and alarm level.
 */
export async function readHeatTagRegisters(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<HeatTagPollResult> {
  const [tempResp, humResp, alarmResp] = await Promise.all([
    client.readHoldingRegisters(REG_HEATTAG_TEMP, 2),
    client.readHoldingRegisters(REG_HEATTAG_HUMIDITY, 2),
    client.readHoldingRegisters(REG_HEATTAG_ALARM, 1),
  ]);

  const humidity = humResp.response.body.valuesAsBuffer.readFloatBE(0);

  return {
    temperature: tempResp.response.body.valuesAsBuffer.readFloatBE(0),
    humidity: humidity * 100, // Convert 0.50 -> 50%
    alarmLevel: alarmResp.response.body.valuesAsBuffer.readUInt16BE(0),
  };
}

/**
 * Read Control 2DI module: two digital input statuses.
 */
export async function readControl2DIRegisters(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<Control2DIPollResult> {
  const [di1Resp, di2Resp] = await Promise.all([
    client.readHoldingRegisters(REG_DI1_STATUS, 1),
    client.readHoldingRegisters(REG_DI2_STATUS, 1),
  ]);

  // Register value: 0=On, 1=Off — invert so true=On
  return {
    di1Status: di1Resp.response.body.valuesAsBuffer.readUInt16BE(0) === 0,
    di2Status: di2Resp.response.body.valuesAsBuffer.readUInt16BE(0) === 0,
  };
}

/**
 * Read Control IO module: digital input + digital output status.
 */
export async function readControlIORegisters(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<ControlIOPollResult> {
  const [di1Resp, doResp] = await Promise.all([
    client.readHoldingRegisters(REG_DI1_STATUS, 1),
    client.readHoldingRegisters(REG_DO1_STATUS, 1),
  ]);

  return {
    di1Status: di1Resp.response.body.valuesAsBuffer.readUInt16BE(0) === 0,
    outputStatus: doResp.response.body.valuesAsBuffer.readUInt16BE(0) === 1,
  };
}

/**
 * Write a command to the Control IO digital output.
 * value true = On (2), value false = Off (1).
 */
export async function writeControlIOOutput(
  client: InstanceType<typeof Modbus.client.TCP>,
  on: boolean,
): Promise<void> {
  await client.writeSingleRegister(REG_DO1_CMD, on ? 2 : 1);
}

/**
 * Read the device type code at a given unit ID.
 * Returns 0 or 65535 if no device is present.
 */
export async function readDeviceType(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<number> {
  const resp = await client.readHoldingRegisters(REG_DEVICE_TYPE, 1);
  return resp.response.body.valuesAsBuffer.readUInt16BE(0);
}

/**
 * Read the user-configured device name (register 31000, 10 regs = 20 ASCII chars).
 * Returns empty string if no name is set.
 */
export async function readDeviceName(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<string> {
  const resp = await client.readHoldingRegisters(REG_DEVICE_NAME, 10);
  return parseAsciiBuffer(resp.response.body.valuesAsBuffer);
}

/**
 * Read the device address table from a PAS600 Panel Server gateway (unit 255).
 *
 * The gateway stores up to 99 device slots. Each slot occupies 5 registers
 * starting at 0x01F8. The first register of each slot holds the Modbus unit ID
 * of the device in that slot (0 or 65535 = empty).
 *
 * Returns a Map of slot number (1-99) → unit ID for occupied slots.
 */
export async function readPanelServerDeviceAddresses(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<Map<number, number>> {
  // 99 slots × 5 registers = 495 registers total, starting at 0x01F8
  // Read in 4 chunks of 125 registers (Modbus max per read)
  const baseReg = REG_PAS_DEVICE_ADDRESS;
  const chunks = await Promise.all([
    client.readHoldingRegisters(baseReg, 125),
    client.readHoldingRegisters(baseReg + 125, 125),
    client.readHoldingRegisters(baseReg + 250, 125),
    client.readHoldingRegisters(baseReg + 375, 120), // last chunk: 495 - 375 = 120
  ]);

  const addresses = new Map<number, number>();

  for (let slot = 1; slot <= 99; slot++) {
    // Each slot is 5 registers; first register = unit ID
    const regOffset = (slot - 1) * 5;
    const chunkIndex = Math.floor(regOffset / 125);
    const indexInChunk = (regOffset % 125) * 2; // 2 bytes per register
    const buf = chunks[chunkIndex].response.body.valuesAsBuffer;
    const unitId = buf.readUInt16BE(indexInChunk);

    if (unitId !== 0 && unitId !== 65535) {
      addresses.set(slot, unitId);
    }
  }

  return addresses;
}

/**
 * Read the commercial reference string from a device (PAS600 Panel Server).
 * Register 0x7954, 16 registers = 32 ASCII chars. Returns e.g. "A9MEM1560".
 */
export async function readCommercialReference(
  client: InstanceType<typeof Modbus.client.TCP>,
): Promise<string> {
  const resp = await client.readHoldingRegisters(REG_COMMERCIAL_REF, 16);
  return parseAsciiBuffer(resp.response.body.valuesAsBuffer);
}

/** Parse a Modbus register buffer as a null-terminated ASCII string. */
function parseAsciiBuffer(buf: Buffer): string {
  let str = '';
  for (let i = 0; i < buf.length; i++) {
    const byte = buf[i];
    if (byte === 0) break;
    str += String.fromCharCode(byte);
  }
  return str.trim();
}
