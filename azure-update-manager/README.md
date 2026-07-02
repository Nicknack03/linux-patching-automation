# Azure Linux VM Patching Automation (Guest Patching API)

A Python automation tool that patches Azure Linux Virtual Machines using the Azure Guest Patching API. The tool performs patch assessment, installs security and critical updates, captures package versions before and after patching, and generates a detailed Excel report.

---

## Features

- Patch multiple Azure Linux VMs in parallel
- Uses Azure Guest Patching API
- Azure Run Command integration for package version snapshots
- Automatic patch assessment
- Install only selected patch classifications
- Configurable reboot behavior
- Tracks installed package version changes
- Parallel execution using ThreadPoolExecutor
- Thread-safe logging
- Comprehensive Excel reporting
- Supports Ubuntu, Debian, RHEL, CentOS, Rocky, AlmaLinux, Oracle Linux, SUSE, and openSUSE

---

## Supported Linux Distributions

| Distribution | Package Manager |
|--------------|----------------|
| Ubuntu | apt |
| Debian | apt |
| RHEL | dnf / yum |
| CentOS | yum |
| SUSE Linux Enterprise | zypper |
| openSUSE | zypper |

---

## Prerequisites

- Python 3.9+
- Azure Subscription
- Azure Linux Virtual Machines
- Azure Service Principal
- Azure Compute API permissions
- Virtual Machine Contributor (or Contributor) role on target VMs
- Azure Run Command enabled on target VMs

---

## Installation

Clone the repository

```bash
git clone https://github.com/yourusername/linux-patching-automation.git

cd azure-linux-guest-patching
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

Configure the following environment variables.

### Linux

```bash
export AZURE_TENANT_ID=xxxxxxxx

export AZURE_CLIENT_ID=xxxxxxxx

export AZURE_CLIENT_SECRET=xxxxxxxx

export AZURE_SUBSCRIPTION_ID=xxxxxxxx
```

---

## Configure Target Virtual Machines

Edit the `PatchConfig` class.

```python
TARGETS = [
    ("ResourceGroup1", "vm_name1"),
    ("ResourceGroup2", "vm_name2"),
]
```

---

## Configuration

| Parameter | Description |
|-----------|-------------|
| TARGETS | Azure Resource Group and VM names |
| PACKAGE_NAME_MASKS | Packages eligible for patching (`["*"]` patches all packages) |
| CLASSIFICATIONS | Patch classifications (Critical, Security, etc.) |
| MAX_PATCH_DURATION | Maximum patch execution time |
| REBOOT_SETTING | Never, IfRequired, or Always |
| MAX_WAIT_SECONDS | Maximum wait time for Azure long-running operations |
| MAX_WORKERS | Number of parallel patch jobs |
| REPORT_DIR | Output folder for reports |

---

## Running the Script

```bash
python patch.py
```

---

## Workflow

1. Authenticate using Azure Service Principal
2. Validate VM connectivity
3. Capture installed package versions (before patching)
4. Assess available patches
5. Install selected patches
6. Capture installed package versions (after patching)
7. Compare package versions
8. Generate Excel report

---

## Excel Report

The script generates an Excel report under the `reports/` directory.

Example:

reports/
patch_report_20260702_104520.xlsx

### Report Sheets

#### 1. Patch Summary

Contains:

- Resource Group
- VM Name
- VM Size
- Location
- Status
- Patch Counts
- Reboot Status
- Package Statistics
- Duration
- Errors

---

#### 2. Run Info

Overall execution summary:

- Report Generation Time
- Total VMs
- Successful VMs
- Failed VMs
- Patch Classifications
- Reboot Setting
- Parallel Workers

---

#### 3. Patch Details

Shows every patch processed for each VM.

Columns include:

- Patch Name
- Catalog Version
- Installed Version Before
- Installed Version After
- KB ID
- Classification
- Installation State

---

#### 4. Package Version Changes

Lists every package whose version changed during patching.

Shows:

- Package Name
- Version Before
- Version After
- Status (Updated, Added, Removed)

---

#### 5. VM Summary Log

Human-readable summary for every VM including:

- VM information
- Patch statistics
- Package version changes
- Errors
- Execution duration

---

## Logging

Example output

2026-07-02 10:30:10 [INFO] [UbuntuVM01] Checking VM connectivity

2026-07-02 10:30:45 [INFO] Assessment complete

2026-07-02 10:33:20 [INFO] Install complete

2026-07-02 10:34:12 [INFO] Captured package versions

2026-07-02 10:34:25 [INFO] Report generated

---

## Patch Configuration Example

Patch all packages

```python
PACKAGE_NAME_MASKS = ["*"]
```

Install only Critical and Security updates

```python
CLASSIFICATIONS = [
    "Critical",
    "Security"
]
```

Always reboot

```python
REBOOT_SETTING = "Always"
```

Reboot only if required

```python
REBOOT_SETTING = "IfRequired"
```

Never reboot

```python
REBOOT_SETTING = "Never"
```

---

## License

MIT License

---

## Author

Azure Linux VM Patching Automation

Built with Python, Azure SDK, and OpenPyXL