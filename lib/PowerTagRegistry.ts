'use strict';

import { PowerTagModelConfig } from './types';

/**
 * Map from device type ID (read from Modbus register 31024) to model configuration.
 *
 * Key differences between models:
 * - phaseCount: 1P models expose only L1 capabilities; 3P models expose L1/L2/L3
 * - voltageMode: Most models read L-N voltage (registers 3027+).
 *   A9MEM1540 and A9MEM1570 read L-L voltage (registers 3019+).
 */
export const POWERTAG_MODELS: ReadonlyMap<number, PowerTagModelConfig> = new Map([
  // M63 series
  [41, { typeId: 41, model: 'A9MEM1520', name: 'PowerTag M63 1P',            phases: '1P',   voltageMode: 'L-N', phaseCount: 1 }],
  [42, { typeId: 42, model: 'A9MEM1521', name: 'PowerTag M63 1P+N Top',      phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [43, { typeId: 43, model: 'A9MEM1522', name: 'PowerTag M63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [44, { typeId: 44, model: 'A9MEM1540', name: 'PowerTag M63 3P',            phases: '3P',   voltageMode: 'L-L', phaseCount: 3 }],
  [45, { typeId: 45, model: 'A9MEM1541', name: 'PowerTag M63 3P+N Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
  [46, { typeId: 46, model: 'A9MEM1542', name: 'PowerTag M63 3P+N Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],

  // F63 / P63 series
  [81, { typeId: 81, model: 'A9MEM1560', name: 'PowerTag F63 1P+N',          phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [82, { typeId: 82, model: 'A9MEM1561', name: 'PowerTag P63 1P+N Top',      phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [83, { typeId: 83, model: 'A9MEM1562', name: 'PowerTag P63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [84, { typeId: 84, model: 'A9MEM1563', name: 'PowerTag P63 1P+N Bottom',   phases: '1P+N', voltageMode: 'L-N', phaseCount: 1 }],
  [85, { typeId: 85, model: 'A9MEM1570', name: 'PowerTag F63 3P+N',          phases: '3P+N', voltageMode: 'L-L', phaseCount: 3 }],
  [86, { typeId: 86, model: 'A9MEM1571', name: 'PowerTag P63 3P+N Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
  [87, { typeId: 87, model: 'A9MEM1572', name: 'PowerTag P63 3P+N Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],

  // NSX series
  [92, { typeId: 92, model: 'LV434020', name: 'PowerTag NSX 250A Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
  [93, { typeId: 93, model: 'LV434021', name: 'PowerTag NSX 250A Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
  [94, { typeId: 94, model: 'LV434022', name: 'PowerTag NSX 630A Top',      phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
  [95, { typeId: 95, model: 'LV434023', name: 'PowerTag NSX 630A Bottom',   phases: '3P+N', voltageMode: 'L-N', phaseCount: 3 }],
]);

/**
 * Build the Homey capability list for a given model.
 * 1P models get only .l1 sub-capabilities; 3P models get .l1/.l2/.l3.
 */
export function getCapabilitiesForModel(config: PowerTagModelConfig): string[] {
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

/**
 * Build capabilitiesOptions for Homey (titles for sub-capabilities).
 */
export function getCapabilityOptionsForModel(config: PowerTagModelConfig): Record<string, object> {
  const opts: Record<string, object> = {};
  const phases = config.phaseCount === 1 ? ['l1'] : ['l1', 'l2', 'l3'];
  const labels: Record<string, string> = { l1: 'L1', l2: 'L2', l3: 'L3' };

  for (const phase of phases) {
    opts[`measure_voltage.${phase}`] = { title: { en: `Voltage ${labels[phase]}` } };
    opts[`measure_current.${phase}`] = { title: { en: `Current ${labels[phase]}` } };
    opts[`measure_power.${phase}`]   = { title: { en: `Power ${labels[phase]}` } };
  }

  return opts;
}
