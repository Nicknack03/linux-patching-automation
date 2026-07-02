# Azure Linux VM Patching Automation using Run Command

A Python automation tool to patch multiple Azure Linux Virtual Machines using **Azure Run Command**. The script connects to Azure using a Service Principal, detects the Linux distribution, installs available updates, optionally reboots the VM if required, and generates a detailed Excel report.

---

## Features

- Patch multiple Azure Linux VMs in parallel
- Uses Azure Run Command (no SSH required)
- Automatic Linux distribution detection
- Supports Ubuntu, Debian, RHEL, CentOS, SUSE, and openSUSE
- Package assessment before installation
- Optional Assess-Only (Dry Run) mode
- Automatic reboot when required
- Package version comparison before and after patching
- Multi-threaded execution
- Detailed Excel reporting
- Thread-safe logging
- Uses Azure Service Principal authentication

---

## Supported Operating Systems

| Distribution | Package Manager |
|--------------|-----------------|
| Ubuntu | apt |
| Debian | apt |
| RHEL | dnf / yum |
| CentOS | yum |
| SUSE Linux Enterprise | zypper |
| openSUSE | zypper |

---

## Prerequisites

- Python 3.9 or later
- Azure Subscription
- Azure Linux Virtual Machines
- Azure Service Principal
- Azure Run Command enabled
- Contributor or Virtual Machine Contributor permissions on the target VMs

---

## Installation

Clone the repository

```bash
git clone https://github.com/yourusername/linux-patching-automation.git

cd azure-linux-patching
```

Create a virtual environment



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

## Azure Authentication

Set the following environment variables.

### Linux

```bash
export AZURE_TENANT_ID=xxxxxxxx

export AZURE_CLIENT_ID=xxxxxxxx

export AZURE_CLIENT_SECRET=xxxxxxxx

export AZURE_SUBSCRIPTION_ID=xxxxxxxx
```

---

## Configure Target VMs

Update the `PatchConfig` class.

```python
TARGETS = [
    ("ResourceGroup1", "vm-name1"),
    ("ResourceGroup2", "vm-name2")
]
```

---

## Configuration Options

| Setting | Description |
|----------|-------------|
| TARGETS | List of Azure Resource Group and VM names |
| ASSESS_ONLY | Only assess updates without installing |
| REBOOT_SETTING | never or if_needed |
| MAX_WORKERS | Number of VMs patched simultaneously |
| RUN_COMMAND_TIMEOUT_SECONDS | Maximum execution time |
| REPORT_DIR | Output directory for reports |

---

## Running the Script

```bash
python patch.py
```

---

## Workflow

1. Authenticate using Azure Service Principal
2. Create Azure Compute Management Client
3. Detect Linux distribution
4. Capture installed package versions
5. Assess available updates
6. Install updates
7. Detect reboot requirement
8. Restart VM if configured
9. Capture package versions after patching
10. Compare package versions
11. Generate Excel report

---

## Generated Excel Report

The script creates an Excel report inside the **reports/** folder.

Example

reports/
linux_runcmd_patch_report_20260702_103025.xlsx

### Worksheet 1 - Results

Contains

- Resource Group
- VM Name
- Operating System
- Version
- Package Manager
- Success / Failure
- Stage Reached
- Reboot Status
- Packages Changed
- Errors
- Duration

---

### Worksheet 2 - Version Changes

Lists

- Updated packages
- Added packages
- Removed packages

---

### Worksheet 3 - VM Detail Log

Contains

- Assessment output
- Installation output
- Summary
- Errors
- Package statistics

---

### Worksheet 4 - Run Info

Displays

- Total VMs
- Successful patch count
- Failed patch count
- Assess Only status
- Reboot configuration
- Parallel worker count

---

## Logging

Example

2026-07-02 10:05:12 [INFO] [UbuntuVM] Detecting distro family

2026-07-02 10:05:30 [INFO] Assessment complete

2026-07-02 10:06:50 [INFO] Install complete

2026-07-02 10:07:15 [INFO] Reboot required

---

## Assess Only Mode

To only check available updates

```python
ASSESS_ONLY = True
```

---

## Reboot Settings

Never reboot

```python
REBOOT_SETTING = "never"
```

Reboot only if required

```python
REBOOT_SETTING = "if_needed"
```

---

## Security Recommendations

- Never hardcode Azure credentials.
- Store credentials in environment variables.
- Prefer Azure Key Vault for production environments.
- Grant only the minimum required Azure RBAC permissions.
- Rotate Service Principal secrets regularly.

---

## Dependencies

- Azure Identity SDK
- Azure Compute Management SDK
- OpenPyXL
- Python Standard Library

---

## License

MIT License

---

## Author

Azure Linux VM Patching Automation

Built with Python, Azure SDK, and OpenPyXL