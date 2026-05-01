# Cisco Tactical Controller — CrayonEaters Series

A PyQt5 GUI application that automates IOS-XE firmware upgrades across a fleet
of Cisco devices using a combination of RESTCONF (for version retrieval), SSH/CLI
(for SCP enablement, flash checks, install commands), and SCP (for image transfer).

---

## Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Installation](#installation)
4. [Project Structure](#project-structure)
5. [Excel Tracker Setup](#excel-tracker-setup)
6. [IOS Repository Setup](#ios-repository-setup)
7. [Running the Application with Python](#running-the-application-with-python)
8. [Workflow Reference](#workflow-reference)
9. [Logging](#logging)
10. [Bugs & Issues Found](#bugs--issues-found)

---

## Overview

The tool reads a device inventory from an Excel workbook, connects to each device
over SSH, retrieves the running IOS version via RESTCONF, compares it against a 
recommended version stored in the sheet, and orchestrates the full install-mode upgrade 
sequence:

```
SSH check → RESTCONF check → Version retrieval → Flash check →
SCP transfer → install add → install activate → (reload) →
install commit → install remove inactive → Excel update
```

Everything is tracked live in the Excel file so you can inspect results per device
at any point during or after the run.

---

## Prerequisites

### Python

Python **3.9 or later** is required.

### Python packages

Install all dependencies with:

```bash
pip install -r requirements.txt
```

Create a `requirements.txt` containing:

```
pandas>=3.0.2
PyQt5>=5.15.11
psutil>=7.2.2
openpyxl>=3.1.5
netmiko>=4.6.0
paramiko>=4.0.0
scp>=0.15.0
aiohttp>=3.13.5
```

### Network requirements

| Protocol | Port | Direction            | Purpose                              |
|----------|------|----------------------|--------------------------------------|
| SSH      | 22   | Controller → Device  | CLI, SCP transfer                    |
| HTTPS    | 443  | Controller → Device  | RESTCONF version retrieval           |
| OOBM     | —    | Out-of-band mgmt     | All connections go via OOBM IP       |

SSL certificate validation is disabled (`ssl=False`) — use only on a trusted
management network.

### Device requirements

- Cisco IOS-XE devices in **install mode** (not bundle mode).
- SSH enabled and reachable on the OOBM IP. 
- RESTCONF and SCP enabled.
- Credentials must have privilege level 15 (or equivalent) to run `install` commands.

---

## Installation

```bash
git clone <repo-url>
cd <repo-name>
pip install -r requirements.txt
```

Place the Excel tracker (`LAN_Network_Devices.xlsx`) and the `ios_repository/`
folder in the same directory as the Python files. See the sections below for
their required formats.

---

## Project Structure

```
cisco-tactical-controller/
│
├── program_gui.py                  # Entry point — PyQt5 main window and workers
├── excel_and_data_ops.py           # All DataFrame and Excel I/O operations
├── device_cli_ops.py               # Netmiko SSH/CLI operations (install, cleanup)
├── device_api_ops.py               # aiohttp RESTCONF operations (version retrieval, polling)
├── device_file_transfer_ops.py     # Paramiko + SCP file transfer
├── ip_addressing_ops.py            # IPv4 address validation helper
├── common_helper_functions.py      # Path resolution and version string normalisation
│
├── LAN_Network_Devices.xlsx        # Device inventory and tracking spreadsheet
│
├── ios_repository/                 # Local IOS image storage (you must create this)
│   ├── C9200/                      # One subfolder per device model (matches "Model" column)
│   │   └── cat9k_lite_iosxe.17.15.04c.SPA.bin
│   ├── C9300/
│   │   └── cat9k_iosxe.17.15.04c.SPA.bin
│   └── ...
│
├── execution_logs.log              # Rotating log file (auto-created on first run)
└── requirements.txt
```

---

## Excel Tracker Setup

The workbook must contain a sheet whose name you enter in the GUI. The sheet
must have **all of the following columns** (order does not matter, but names
must match exactly):

| Column                            | Type     | Populated by    | Description                                                |
|-----------------------------------|----------|-----------------|------------------------------------------------------------|
| `Hostname`                        | Static   | You             | Device hostname, used as the lookup key                    |
| `Model`                           | Static   | You             | Must match the subfolder name in `ios_repository/`         |
| `OOBM IP Address`                 | Static   | You             | Management IP used for all connections                     |
| `Recommended IOS Version`         | Static   | You             | Target version string, e.g. `17.15.4c`                     |
| `Current IOS Version`             | Dynamic  | Tool            | Retrieved live via RESTCONF                                |
| `Needs Update`                    | Dynamic  | Tool            | `YES` / `NO` / `UNKNOWN`                                   |
| `Update IOS File Present`         | Dynamic  | Tool            | `YES` / `NO` / `N/A`                                       |
| `Enough Flash Space`              | Dynamic  | Tool            | `YES` / `NO` / `UNKNOWN` / `N/A`                           |
| `Status`                          | Dynamic  | Tool            | `ONLINE` / `OFFLINE`                                       |
| `Auth Status`                     | Dynamic  | Tool            | `AUTH_OK` / `AUTH_BAD` / `N/A`                             |
| `Transfer Result`                 | Dynamic  | Tool            | `SUCCESS` / `FAILED` / `N/A`                               |
| `Install Status`                  | Dynamic  | Tool            | `INSTALL_TRIGGERED` / `ADD_FAILED` / `NOT_ATTEMPTED` / …   |
| `Update Result`                   | Dynamic  | Tool            | `SUCCESS` / `FAILED` / `COMMIT_FAILED` / `UNKNOWN` / `N/A` |
| `Cleaned Inactive`                | Dynamic  | Tool            | `CLEANED` / `NOTHING_TO_CLEAN` / `N/A`                     |

All **Dynamic** columns are wiped and rewritten on every run. Do not put
manual data in them.

> **Close the Excel file before running.** The tool checks for open file handles
> at startup and will abort with an error if the file is locked.

---

## IOS Repository Setup

Then `ios_repository/` folder is in the application directory. Inside it,
create one subfolder per device model. Some of them already exist. The subfolder name must exactly match
the value in the `Model` column of your Excel sheet.

Place the `.bin` IOS image file inside the matching model folder. The filename
must contain the version string from the `Recommended IOS Version` column.
Leading zeros in version segments are handled automatically
(e.g. `17.15.04c` and `17.15.4c` are treated as equal).

```
ios_repository/
└── C9300/
    └── cat9k_iosxe.17.15.04c.SPA.bin   ← filename must contain "17.15.4c" or "17.15.04c"
```

---

## Running the Application with Python

```bash
python program_gui.py
```

### Step-by-step GUI walkthrough
<br/>
1. Enter the Excel sheet name in the text field at the top (e.g. `LAN`).
<br/>
<br/>
2. Click `SHOW ELIGIBLE DEVICES`.
<br/>
<br/>
3. Enter device credentials in the popup (username + password).  
   The tool connects to all devices over SSH,
   checks versions and flash space, then populates the Excel tracker.
<br/>
<br/>
4. Review the table. Only devices that pass all eligibility checks appear:
   ONLINE, AUTH_OK, RESTCONF OPERATIVE, IOS image present locally, enough
   flash space, and current version differs from recommended.
<br/>
<br/>
5. Select devices using the checkboxes (or click `ALL`).
<br/>
<br/>
6. Click `START UPDATE`.
<br/>
<br/>
7. Choose the transfer mode:
   - Sequential — one device at a time; safe, readable progress per device.
   - Threaded — all devices simultaneously; faster but progress output overlaps.
<br/>
<br/>
8. The tool runs the full install sequence autonomously. A summary dialog appears
   on completion. The Excel tracker is updated after each stage.

### Cancel during transfer

The `CANCEL UPDATE` button is active **only during the SCP transfer stage**.
Clicking it:
- Sets a global cancellation flag to prevent new transfers from starting.
- Connects to each device and clears all VTY lines (ending the SCP session).
- Deletes the partially transferred file from `bootflash:`.

---

## Workflow Reference

| Stage | Function | Notes |
|-------|----------|-------|
| Enable APIs | `enable_scp_and_restconf_all` | Runs `restconf`, `ip scp server enable`, `ip http secure-server`. **Not saved** to startup-config. |
| RESTCONF poll | `get_all_restconf_status` | Polls `/restconf` endpoint; 180 s timeout during Show, 720 s after reload. |
| Version retrieval | `get_all_versions` | Uses `Cisco-IOS-XE-device-hardware-oper` YANG model. |
| Flash check | `get_flash_free_space_all` | Parses `show file systems`; prefers `bootflash:` over `flash:`. |
| SCP transfer | `threaded_transfer_ios_image_all` | Transfers to `bootflash:/`. Progress bar per device in terminal. |
| Boot variable | `install_ios_image` (stage 1) | Pushes `boot system bootflash:packages.conf` and saves config before install commands. |
| Install add | `install_ios_image` (stage 2) | 900 s read timeout. Detects dead peers via TCP keepalives + watchdog thread. |
| Install activate | `install_ios_image` (stage 3) | Uses `send_multiline` to handle the `[y/n]` reload confirmation. |
| Post-reload wait | `populate_post_install_columns` | Sleeps 6 minutes, then polls RESTCONF for up to 12 minutes. |
| Install commit | `commit_ios_install_all` | Runs `install commit` to persist the new image. |
| Version verify | `get_all_versions` (post-install) | Confirms running version matches recommended. |
| Cleanup | `remove_inactive_ios_all` | Runs `install remove inactive` and auto-confirms. It also disables RESTCONF and HTTPS server.|

---

## Logging

All operations are logged to `execution_logs.log` in the application directory.
The file rotates at **5 MB** and is replaced (not backed up) when full.

Log format:
```
2025-01-15 14:32:01,234 - device_cli_ops - INFO - 'SW-CORE-01' (10.1.1.1): boot system commands pushed and config saved
```

---

## Bugs & Issues Found

The following bugs and issues have been identified and are still persistent on the code.

---

### 🔴 BUG 0001
When the SCP transfer is finished, there are a couple commands that need to be pushed to the devices. If a device goes offline
during this commands push, the whole program will freeze for 15 minutes (the 'read_timeout' value) before raising an error and
notifying the user via the GUI. The Paramiko TCP socket will be closed in time (just 70 seconds after the loss of connectivity)
but Netmiko will still try to read the output, resulting in the GUI getting the error codes after 15 minutes. The error will be notified 
in time in the log file.