# M365 License Usage Report (PowerShell)

This repository contains a PowerShell script that builds a Microsoft 365 license utilization report by correlating assigned licenses with actual workload usage signals.

## What it checks

For each licensed user, the script aggregates all assigned SKUs, inspects the combined service plans, and checks usage signals:

- **Intune-related plans** (`INTUNE`, `EMS`, `AAD_PREMIUM`): user has one or more Intune managed devices.
- **Exchange plans**: mailbox activity in the selected reporting period.
- **OneDrive/SharePoint plans**: OneDrive account activity in the selected reporting period.
- **Teams plans**: Teams activity in the selected reporting period.

A utilization state is computed per user:

- `Used`
- `PartiallyUsed`
- `Unused`
- `NoTrackedWorkload`

## Authentication model

The script uses **interactive delegated authentication** with `Connect-MgGraph`, prompting you to sign in with an admin account. It does **not** require app registration.

## Requirements

- PowerShell 7+ recommended.
- Ability to install PowerShell modules from PSGallery (unless already installed).
- Admin account with permissions to consent/access:
  - `User.Read.All`
  - `Directory.Read.All`
  - `DeviceManagementManagedDevices.Read.All`
  - `Reports.Read.All`
  - `AuditLog.Read.All`

## Usage

```powershell
./M365-License-Usage-Report.ps1
```

### Common options

```powershell
# Generate HTML (default), 30-day period, output to ./output
./M365-License-Usage-Report.ps1

# Generate CSV only
./M365-License-Usage-Report.ps1 -OutputFormat Csv

# Generate both HTML and CSV for a 90-day period
./M365-License-Usage-Report.ps1 -OutputFormat Both -Period D90

# Custom output directory
./M365-License-Usage-Report.ps1 -OutputFolder "C:\Reports\M365"

# Skip auto module install (fails if modules are missing)
./M365-License-Usage-Report.ps1 -SkipModuleInstall
```

## Output

- HTML report: `M365_License_Utilization_<timestamp>.html`
- CSV report: `M365_License_Utilization_<timestamp>.csv`

Both are written to the selected output folder.
