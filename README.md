# Schneider EcoStruxure

Monitor your Schneider EcoStruxure energy devices in Homey via Modbus TCP.

## Supported Gateways

- Smartlink SI D
- PowerTag Link
- EcoStruxure Panel Server

## Supported Devices

- PowerTag M63 series (1P, 1P+N, 3P, 3P+N)
- PowerTag F63 series (1P+N, 3P+N)
- PowerTag P63 series (1P+N, 3P+N)
- PowerTag NSX series (250A, 630A)

These are Schneider EcoStruxure Acti9 PowerTag energy sensors that clip onto your circuit breakers and NSX compact circuit breakers.

## Capabilities

Each PowerTag exposes the following measurements:

- **Power** — total active power (W)
- **Energy** — cumulative energy consumption (kWh)
- **Voltage** — per phase (V)
- **Current** — per phase (A)
- **Power per phase** — per phase (W)
- **Power factor**
- **Frequency** (Hz)
- **Temperature** (°C)

3-phase devices report values for L1, L2, and L3 individually.

## Pairing

1. Enter the IP address and port of your gateway
2. The app scans for connected PowerTag devices via Modbus
3. Select the devices you want to add

## Configuration

Each device has the following settings:

- **Gateway IP Address** — IP of your Smartlink SI D, PowerTag Link, or Panel Server
- **Gateway Port** — Modbus TCP port (default: 502)
- **Polling Interval** — how often to read measurements (default: 30 seconds)
