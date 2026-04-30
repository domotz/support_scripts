# Domotz Inventory Overviewer - Export to Excel Tool

A PowerShell tool designed to extract data from Domotz (Organizations, Collectors, Devices) and export to a formatted Excel file.

The script retrieves information from your Domotz account via the API and creates an Excel workbook with separate worksheets for Organizations, Collectors, and Devices.

---

## System Requirements - Mandatory Requirements

- **Windows PC** - This script only runs on Windows
- **PowerShell** - Pre-installed on Windows
- **Microsoft Excel** - Must be installed on the PC
- **ImportExcel Module** - PowerShell module (installation instructions below)

## First-Time Setup

### Step 1: Execution Policy

Before running any PowerShell script, you need to allow script execution:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

When prompted, respond with:

- `Y` (for Yes) - Allow for this session only
- `A` (for All) - Allow for all scripts in this session

> **Note:** This only affects the current PowerShell session and doesn't change your system-wide settings.

### Step 2: Install ImportExcel Module

**Check if already installed:**

```powershell
Get-Module -ListAvailable -Name ImportExcel
```

**If not installed:**

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

### Step 3: Configure API Access

Create a file named `.env` in the same folder as the script:

```env
API_KEY='your_actual_api_key'
BASE_URL='https://api-us-east-1-cell-1.domotz.com/public-api/v1'
```

**BASE_URL by Region:**

- **US Region:** `https://api-us-east-1-cell-1.domotz.com/public-api/v1`
- **EU Region:** `https://api-eu-west-1-cell-1.domotz.com/public-api/v1`

> **Tip:** Get your API key from Domotz account settings. See [Domotz API Documentation](https://help.domotz.com/admin-global-features/domotz-api/) for details.

## Getting Help

### View Usage Information

To see complete usage instructions and available operations:

```powershell
.\inventory_overviewer.ps1 -help
# or
.\inventory_overviewer.ps1 -?
```

---

## Quick Start

### Extract All Data (Default)

Run the script without parameters to extract all data (Organizations, Collectors, Devices):

```powershell
.\inventory_overviewer.ps1
```

This will:

- Retrieve all collectors from your account
- Count managed devices per collector
- Extract Organizations, Collectors, and Devices data
- Create a formatted Excel file with three worksheets
- Auto-open the file for review

---

## Available Operations

| Operation           | Description                   | Required Parameters | Optional Parameters                               |
| ------------------- | ----------------------------- | ------------------- | ------------------------------------------------- |
| `extract` (default) | Extract data and create Excel | None                | `-collector_ids`<br>`-device-type`<br>`-filename` |
| `list_collectors`   | List all collectors/agents    | None                | None                                              |

---

## Usage Examples

### List All Collectors

```powershell
.\inventory_overviewer.ps1 -operation list_collectors
```

Shows all collectors/agents in your Domotz account with their IDs.

### Extract All Data

```powershell
.\inventory_overviewer.ps1
```

Extracts Organizations, Collectors, and Devices to Excel.

### Extract from Specific Collectors

```powershell
.\inventory_overviewer.ps1 -collector_ids "312189,313759"
```

> **Note:** When using `-collector_ids`, the Organizations worksheet is not created (since you're filtering to specific collectors).

### Extract Only Managed Devices

```powershell
.\inventory_overviewer.ps1 -device-type managed
```

### Extract Only Unmanaged Devices

```powershell
.\inventory_overviewer.ps1 -device-type unmanaged
```

### Custom Output Filename

```powershell
.\inventory_overviewer.ps1 -filename "domotz_export_2025"
```

### Combine Options

```powershell
.\inventory_overviewer.ps1 -collector_ids "312189" -device-type managed -filename "site_a_devices"
```

---

## Parameters

| Parameter        | Description                               | Values                                      | Default                |
| ---------------- | ----------------------------------------- | ------------------------------------------- | ---------------------- |
| `-operation`     | Operation to perform                      | `extract`, `list_collectors`                | `extract`              |
| `-collector_ids` | Filter by collector IDs (comma-separated) | Any valid collector IDs                     | All collectors         |
| `-device-type`   | Filter device types                       | `managed`, `unmanaged`, `managed,unmanaged` | `managed,unmanaged`    |
| `-filename`      | Custom output Excel filename              | Any string                                  | `inventory_overviewer` |
| `-debug`         | Enable detailed logging                   | Switch                                      | Off                    |
| `-help`          | Show help                                 | Switch                                      | -                      |

---

## Understanding the Excel Output

### Worksheets Created

| Worksheet         | Description                                                   | Created When               |
| ----------------- | ------------------------------------------------------------- | -------------------------- |
| **Organizations** | List of organizations with collector counts and device totals | No `-collector_ids` filter |
| **Collectors**    | Detailed collector information                                | Always                     |
| **Devices**       | All devices (managed and/or unmanaged)                        | Always                     |

### Organizations Worksheet Columns

- `organization_id` - Organization ID
- `organization_name` - Organization name
- `collector_count` - Number of collectors in this organization
- `number_of_managed_devices` - Total managed devices across all collectors
- `collector_ids` - List of collector IDs (one per line)
- `collector_names` - List of collector names (one per line)

### Collectors Worksheet Columns

- `collector_id` - Collector/agent ID (hyperlinked to Domotz portal)
- `display_name` - Collector name
- `number_of_managed_devices` - Count of managed devices
- `status` - Online/Offline status
- `organization_id`, `organization_name` - Organization info
- `agent_version`, `package_version` - Software versions
- `licence_code`, `licence_type` - License information
- `bound_mac_address` - MAC address of the collector
- `wan_ip`, `wan_hostname` - WAN connection info
- `latitude`, `longitude` - Location coordinates
- And more...

### Devices Worksheet Columns

- `collector_id`, `collector_name` - Parent collector info
- `device_id` - Device ID
- `display_name` - Device name (hyperlinked to Domotz portal)
- `device_type` - `managed` or `unmanaged`
- `status` - Device status (managed devices only)
- `ip_addresses` - IP address(es)
- `hw_address` - MAC address
- `vendor`, `model` - Device manufacturer info
- `type_label` - Device type (e.g., "Network Equipment", "Audio & Video")
- And more...

### Visual Formatting

The Excel file includes professional formatting:

- **Yellow header row** with bold text
- **AutoFilter** enabled on all columns
- **Borders** on all cells
- **Hyperlinks** on device and collector names (blue, underlined)
- **Unmanaged devices** styled with dark grey text
- **Empty cells** in unmanaged rows have light grey fill
- **Devices worksheet** is the active tab when opened

---

## Troubleshooting

### Script Won't Run - Execution Policy Error

**Error:** `...cannot be loaded because running scripts is disabled...`

**Solution:**

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Respond `Y` or `A` when prompted, then run the script again.

### ImportExcel Module Not Found

**Error:** `The term 'Import-Excel' is not recognized...`

**Solution:**

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

### Invalid Parameter Error

**Error:** `Unknown parameter(s) detected...`

**Solution:**

- Check parameter names for typos
- Use `-help` to see valid parameters
- Valid parameters: `-operation`, `-collector_ids`, `-device-type`, `-filename`, `-debug`, `-help`

### Invalid Collector IDs

**Error:** `None of the specified collector IDs were found...`

**Solution:**

- Run `-operation list_collectors` to see valid IDs
- Verify the collector IDs in your command

### API Key Issues

**Error:** `401: Unauthorized`

**Solution:**

- Verify `.env` file is in the same folder as the script
- Check `API_KEY` is correct (no extra quotes/spaces)
- Ensure `BASE_URL` matches your region
- Verify API key is active in Domotz

---

## Files Generated

- **Excel File:** `inventory_overviewer.xlsx` (or custom name)
- **Log File:** `inventory_overviewer_Log.txt`

> **Note:** If a file already exists, a timestamp is automatically appended to avoid overwriting.

The log file contains complete execution history, API calls, and processing details for troubleshooting.

---

## FAQ

**Q: Can I filter to specific collectors?**  
A: Yes, use `-collector_ids "id1,id2,id3"` to extract data only from those collectors.

**Q: Why is the Organizations worksheet missing?**  
A: When you use `-collector_ids` filter, the Organizations worksheet is skipped since you're filtering to specific collectors.

**Q: What's the difference between managed and unmanaged devices?**  
A: Managed devices are actively monitored by Domotz with full details. Unmanaged devices are detected but not actively monitored, so they have fewer data fields.

**Q: How can I tell managed from unmanaged devices in the Excel?**  
A: The `device_type` column shows "managed" or "unmanaged". Additionally, unmanaged devices are styled with dark grey text and grey fill on empty cells.

**Q: Can I export to CSV instead of Excel?**  
A: The script creates Excel files (.xlsx). You can open the file and save as CSV from Excel if needed.

**Q: How often should I run the extraction?**  
A: As needed. Each run creates a snapshot of your current Domotz data.

---

## License

This script is provided "AS IS" for illustrative/educational purposes. See script header for full disclaimer.
