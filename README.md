# FlowController

Desktop GUI application for real-time monitoring and control of mass flow controllers over serial communication.

*Property of GP Ionics*

---

## Overview

FlowController is a .NET/Avalonia UI desktop application built to interface with Mass flow controllers (MFCs) over serial. It auto-discovers connected devices, polls live telemetry at 250 ms intervals, and lets operators adjust flow setpoints and device IDs directly from the GUI — no terminal needed.

The current version is a full rewrite in C# from an earlier Python prototype (see below). A companion C++ implementation also exists as a reference for understanding the underlying serial communication protocol and control logic at a lower level.

---

## Features

- **Auto-discovery** — scans all available serial ports on startup and detects connected MFC devices by querying IDs A–Z. Rescans automatically when devices are plugged or unplugged (Windows via WMI, Linux via `/dev` watcher)
- **Live telemetry** — polls flow rate, temperature, setpoint, drive %, volumetric total, and units at 4 Hz for the selected device
- **Setpoint control** — arrow key input with acceleration (step size increases the longer you hold), with debounced serial write on key release
- **Unit conversion** — full bidirectional conversion between SCCM, SLPM, SmL/s, SCFM, and more
- **Device ID reassignment** — rename a device's single-letter ID (A–Z) from the Setup menu without reflashing
- **Error indicators** — live display of TOV, MOV, OVR, HLD, and VTM fault flags from the device status frame
- **Cross-platform** — runs on Windows and Linux (including Raspberry Pi)

---

## Stack

| Layer | Technology |
|---|---|
| UI Framework | [Avalonia UI](https://avaloniaui.net/) (.NET) |
| Language | C# |
| Serial Communication | .NET `SerialPort` — 38400 baud, 8N1, ASCII, `\r` terminated |
| Device Watcher | WMI (`Win32_DeviceChangeEvent`) on Windows / `FileSystemWatcher` on Linux |
| Build | .NET SDK / Visual Studio (`FlowController.slnx`) |
| Target Platform | Windows / Linux |

---

## Serial Protocol

Devices are addressed by a single uppercase letter ID (A–Z). Commands follow the Alicat serial format:

| Command | Description |
|---|---|
| `{ID}SN\r` | Query serial number — used to confirm device presence |
| `{ID}\r` | Poll full status frame (flow, temp, setpoint, drive, flags) |
| `{ID}FPF 0\r` | Query flow unit and full-scale range |
| `{ID}FPF 1\r` | Query volumetric unit |
| `{ID}DV 2\r` | Query current setpoint |
| `{ID}S {value}\r` | Set flow setpoint |
| `{ID}@={newID}\r` | Rename device ID |
| `*LSS S\r` | Broadcast scan trigger |

All commands are terminated with `\r`. Responses are read character-by-character until `\r` or `\n` is received. Each command uses a 100–600 ms timeout with a shared per-port semaphore to prevent concurrent access.

---

### Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download) (8.0 or later)
- Serial adapter connected to the MFC
- Windows or Linux (tested on Raspberry Pi)

---

## Previous Versions

### Python Prototype

The original version was a Python application with a Tkinter GUI and manually scripted serial communication. It validated the core protocol and control logic before the C# rewrite.

![Python Version](Python%20Version.jpeg)

### C++ Reference Implementation

A C++ version also exists as a lower-level reference for the serial communication layer and control logic. It was written to understand how to implement the protocol close to the metal, without a framework abstraction layer.

The C#/Avalonia rewrite builds on both, adding a responsive cross-platform UI, automatic device discovery, and a cleaner separation between the serial and interface layers.
