'use strict';

import { PowerTagModelConfig } from './types';

/**
 * Map from device type ID (read from Modbus register 31024) to model configuration.
 *
 * Key differences between models:
 * - deviceCategory: determines which registers to read and capabilities to expose
 * - phaseCount: 1P models expose only L1 capabilities; 3P models expose L1/L2/L3
 * - voltageMode: 3P (no neutral) models read L-L voltage (registers 3019+).
 *   Models with neutral read L-N voltage (registers 3027+).
 */
export const POWERTAG_MODELS: ReadonlyMap<number, PowerTagModelConfig> = new Map([
  // M63 series
  [41, { typeId: 41,  model: 'A9MEM1520', name: 'PowerTag M63 1P',            phases: '1P',   voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [42, { typeId: 42,  model: 'A9MEM1521', name: 'PowerTag M63 1P+N Top',      phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [43, { typeId: 43,  model: 'A9MEM1522', name: 'PowerTag M63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [44, { typeId: 44,  model: 'A9MEM1540', name: 'PowerTag M63 3P',            phases: '3P',   voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],
  [45, { typeId: 45,  model: 'A9MEM1541', name: 'PowerTag M63 3P+N Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [46, { typeId: 46,  model: 'A9MEM1542', name: 'PowerTag M63 3P+N Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [96, { typeId: 96,  model: 'A9MEM1543', name: 'PowerTag M63 3P 230V',       phases: '3P',   voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],

  // F63 / P63 series
  [81, { typeId: 81,  model: 'A9MEM1560', name: 'PowerTag F63 1P+N',          phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [82, { typeId: 82,  model: 'A9MEM1561', name: 'PowerTag P63 1P+N Top',      phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [83, { typeId: 83,  model: 'A9MEM1562', name: 'PowerTag P63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [84, { typeId: 84,  model: 'A9MEM1563', name: 'PowerTag P63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [85, { typeId: 85,  model: 'A9MEM1570', name: 'PowerTag F63 3P+N',          phases: '3P+N', voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],
  [86, { typeId: 86,  model: 'A9MEM1571', name: 'PowerTag P63 3P+N Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [87, { typeId: 87,  model: 'A9MEM1572', name: 'PowerTag P63 3P+N Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [101, { typeId: 101, model: 'A9MEM1564', name: 'PowerTag F63 1P+N 110V',    phases: '1P+N', voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'energy' }],
  [102, { typeId: 102, model: 'A9MEM1573', name: 'PowerTag F63 3P',           phases: '3P',   voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],
  [103, { typeId: 103, model: 'A9MEM1574', name: 'PowerTag F63 3P+N 110/230V', phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],

  // NSX series (M250/M630)
  [92, { typeId: 92,  model: 'LV434020', name: 'PowerTag NSX 250A 3P',        phases: '3P',   voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],
  [93, { typeId: 93,  model: 'LV434021', name: 'PowerTag NSX 250A 4P',        phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [94, { typeId: 94,  model: 'LV434022', name: 'PowerTag NSX 630A 3P',        phases: '3P',   voltageMode: 'L-L', phaseCount: 3, deviceCategory: 'energy' }],
  [95, { typeId: 95,  model: 'LV434023', name: 'PowerTag NSX 630A 4P',        phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],

  // Rope series
  [104, { typeId: 104, model: 'A9MEM1590', name: 'PowerTag Rope 200A',        phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [105, { typeId: 105, model: 'A9MEM1591', name: 'PowerTag Rope 600A',        phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [106, { typeId: 106, model: 'A9MEM1592', name: 'PowerTag Rope 1000A',       phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],
  [107, { typeId: 107, model: 'A9MEM1593', name: 'PowerTag Rope 2000A',       phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],

  // F160 series
  [121, { typeId: 121, model: 'A9MEM1580', name: 'PowerTag F160',             phases: '3P+N', voltageMode: 'L-N', phaseCount: 3, deviceCategory: 'energy' }],

  // Control modules
  [97, { typeId: 97,  model: 'A9XMC2D3', name: 'PowerTag C 2DI',             phases: '1P',   voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'control_2di' }],
  [98, { typeId: 98,  model: 'A9XMC1D3', name: 'PowerTag C IO',              phases: '1P',   voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'control_io' }],

  // HeatTag sensor
  [171, { typeId: 171, model: 'SMT10020', name: 'HeatTag Sensor',            phases: '1P',   voltageMode: 'L-N', phaseCount: 1, deviceCategory: 'heattag' }],
]);

/**
 * Look up a model config by its commercial reference string (e.g. "A9MEM1560").
 * Used by PAS600 Panel Server discovery where the device type register (31024)
 * is not available and we identify devices by their reference string instead.
 */
export function getModelByReference(ref: string): PowerTagModelConfig | undefined {
  const trimmed = ref.trim();
  for (const config of POWERTAG_MODELS.values()) {
    if (config.model === trimmed) return config;
  }
  return undefined;
}

/**
 * Build the Homey capability list for a given model.
 */
export function getCapabilitiesForModel(config: PowerTagModelConfig): string[] {
  switch (config.deviceCategory) {
    case 'heattag':
      return ['measure_temperature', 'measure_humidity', 'alarm_heat'];

    case 'control_2di':
      return ['alarm_contact.di1', 'alarm_contact.di2'];

    case 'control_io':
      return ['alarm_contact.di1', 'onoff'];

    case 'energy':
    default: {
      const caps: string[] = [
        'measure_power',
        'meter_power',
        'measure_power_factor',
        'measure_temperature',
        'measure_frequency',
      ];

      const phases = config.phaseCount === 1 ? ['l1'] : ['l1', 'l2', 'l3'];
      for (const phase of phases) {
        caps.push(`measure_voltage.${phase}`);
        caps.push(`measure_current.${phase}`);
        caps.push(`measure_power.${phase}`);
      }

      return caps;
    }
  }
}

/**
 * Build capabilitiesOptions for Homey (titles for sub-capabilities).
 */
export function getCapabilityOptionsForModel(config: PowerTagModelConfig): Record<string, object> {
  const opts: Record<string, object> = {};

  switch (config.deviceCategory) {
    case 'control_2di':
      opts['alarm_contact.di1'] = { title: { en: 'Digital Input 1' } };
      opts['alarm_contact.di2'] = { title: { en: 'Digital Input 2' } };
      break;

    case 'control_io':
      opts['alarm_contact.di1'] = { title: { en: 'Digital Input' } };
      break;

    case 'energy':
    default: {
      const phases = config.phaseCount === 1 ? ['l1'] : ['l1', 'l2', 'l3'];
      const labels: Record<string, string> = { l1: 'L1', l2: 'L2', l3: 'L3' };

      for (const phase of phases) {
        opts[`measure_voltage.${phase}`] = { title: { en: `Voltage ${labels[phase]}` } };
        opts[`measure_current.${phase}`] = { title: { en: `Current ${labels[phase]}` } };
        opts[`measure_power.${phase}`]   = { title: { en: `Power ${labels[phase]}` } };
      }
      break;
    }
  }

  return opts;
}
