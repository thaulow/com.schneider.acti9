'use strict';

/** Phases this PowerTag model measures */
export type PhaseConfig = '1P' | '1P+N' | '3P' | '3P+N';

/** How voltage is measured — Line-to-Neutral or Line-to-Line */
export type VoltageMode = 'L-N' | 'L-L';

/** Device category determines which registers to read and which capabilities to expose */
export type DeviceCategory = 'energy' | 'heattag' | 'control_2di' | 'control_io';

/** Configuration for a single PowerTag model */
export interface PowerTagModelConfig {
  typeId: number;
  model: string;
  name: string;
  phases: PhaseConfig;
  voltageMode: VoltageMode;
  phaseCount: 1 | 3;
  deviceCategory: DeviceCategory;
}

/** Data returned from polling an energy sensor */
export interface EnergyPollResult {
  currentL1: number;
  currentL2: number;
  currentL3: number;
  voltagePh1: number;
  voltagePh2: number;
  voltagePh3: number;
  powerL1: number;
  powerL2: number;
  powerL3: number;
  totalPower: number;
  powerFactor: number;
  totalEnergy: number; // kWh (converted from Wh)
  temperature: number;
  frequency: number;
}

/** Data returned from polling a HeatTag sensor */
export interface HeatTagPollResult {
  temperature: number;
  humidity: number;
  alarmLevel: number; // 0=None, 1=Low, 2=Medium, 3=High
}

/** Data returned from polling a Control 2DI module */
export interface Control2DIPollResult {
  di1Status: boolean;
  di2Status: boolean;
}

/** Data returned from polling a Control IO module */
export interface ControlIOPollResult {
  di1Status: boolean;
  outputStatus: boolean;
}

/** Union of all poll result types */
export type PollResult = EnergyPollResult | HeatTagPollResult | Control2DIPollResult | ControlIOPollResult;

/** Settings stored per device */
export interface PowerTagSettings {
  address: string;
  port: number;
  polling: number;
}

/** Store values persisted per device */
export interface PowerTagStore {
  unitId: number;
  typeId: number;
  model: string;
}

/** Device data (unique identifier for Homey) */
export interface PowerTagDeviceData {
  id: string; // "{ip}:{port}:{unitId}"
}
