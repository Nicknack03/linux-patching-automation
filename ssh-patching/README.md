# Linux SSH Patching Automation

A Python-based Linux patching automation tool that securely patches multiple Linux virtual machines over SSH using Paramiko.

The script automatically detects the Linux distribution, performs patch assessment, installs available updates, optionally reboots the system if required, captures package version changes, and generates a detailed Excel report.

---

## Features

- SSH-based patching
- Supports multiple Linux distributions
  - Ubuntu
  - Debian
  - RHEL
  - CentOS
  - SUSE Linux Enterprise
  - openSUSE
- Automatic OS detection
- Package assessment before installation
- Package installation
- Optional automatic reboot
- Parallel patching of multiple VMs
- Package version comparison
- Excel report generation
- Logging with timestamps
- Thread-safe execution

---

## Supported Package Managers

| Distribution | Package Manager |
|--------------|----------------|
| Ubuntu / Debian | apt |
| RHEL / CentOS   | dnf / yum |
| SUSE / openSUSE | zypper |

---

## Installation

Clone the repository

```bash
git clone https://github.com/yourusername/linux-ssh-patching.git

cd linux-ssh-patching
```

Create a virtual environment

---

### Linux

```bash
python3 -m venv .venv

source .venv/bin/activate
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

## Configuration

Edit the `PatchConfig` class.

Example:

```python
TARGETS = [
    {
        "vm_name": "vm-name",
        "host": "host-ip",
        "port": 22,
        "username": "user-name",
        "password": "password"
    }
]
```

---

## Configuration Options

| Setting | Description |
|----------|-------------|
| TARGETS | List of Linux servers |
| ASSESS_ONLY | Only check updates without installing |
| REBOOT_SETTING | `never` or `if_needed` |
| MAX_WORKERS | Number of parallel patching threads |
| REPORT_DIR | Directory for generated reports |

---

## Running the Script

```bash
python patch.py
```

---

## Workflow

1. Connect to VM using SSH
2. Detect Linux distribution
3. Capture installed package versions
4. Check available updates
5. Install updates
6. Detect reboot requirement
7. Reboot if configured
8. Capture updated package versions
9. Compare versions
10. Generate Excel report

---

## Generated Reports

The script generates an Excel report inside the **reports/** directory.

Example:

reports/
    linux_ssh_patch_report_20260702_101530.xlsx

The report contains four worksheets:

### Worksheet 1 - Results

- VM Name
- Host
- Operating System
- Version
- Package Manager
- Success/Failure
- Reboot Status
- Packages Changed
- Duration

### Worksheet 2 - Version Changes

Lists:

- Updated packages
- Added packages
- Removed packages

### Worksheet 3 - VM Detail Log

Contains

- Assessment output
- Installation output
- Summary
- Errors
- Package statistics

### Worksheet 4 - Run Info

Overall execution summary including

- Total VMs
- Success count
- Failure count
- Assess-only mode
- Parallel worker count

---

## Logging

Example

2026-07-02 10:15:33 [INFO] [Ubuntu-01] SSH connected
2026-07-02 10:15:36 [INFO] Assessment complete
2026-07-02 10:16:20 [INFO] Install complete
2026-07-02 10:17:10 [INFO] Reboot required

---

## Supported Reboot Modes

### Never reboot

```python
REBOOT_SETTING = "never"
```

### Reboot only if required

```python
REBOOT_SETTING = "if_needed"
```

---

## Assess Only Mode

To only check available updates without installing them:

```python
ASSESS_ONLY = True
```

---

## Requirements

- Python 3.9+
- SSH access to target VMs
- sudo privileges
- Internet access from target VMs for package repositories

---

## License

MIT License

---

## Author

Linux SSH Patching Automation
Built with Python, Paramiko, and OpenPyXL.