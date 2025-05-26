# Domotz Custom Driver Mass Application Guide

This guide will help you efficiently apply one or more custom drivers to multiple devices in your Domotz account using PowerShell scripts. Follow the workflow below to ensure a smooth process.

## **Prerequisites**
1. **Ensure the following files are in the same directory:**
   - `Devices.csv` → The CSV file where you will define device associations.
   - `Device-List-Script.ps1` → Script to retrieve your Domotz device inventory.
   - `listCustomDrivers.ps1` → Script to retrieve all custom drivers and their IDs.
   - `MassApplyScripts.ps1` → Script to apply custom drivers in bulk.

2. **Ensure PowerShell Execution Policy allows script execution.**  
   If needed, run the following in PowerShell (as Administrator):  
   ```powershell
   Set-ExecutionPolicy Unrestricted -Scope Process

Workflow Steps
Step 1: Retrieve Your Device Inventory
	Run the Device-List-Script.ps1 script to generate a CSV file containing all the devices in your Domotz account. This file will include important information such as:
		- Device ID
		- Agent ID
		- Device Name
	Use this list to identify the devices you want to apply custom drivers to.

Step 2: Populate Devices.csv with Devices & Credentials
	1. Copy and paste the relevant devices from the device inventory CSV into Devices.csv.
	2. Fill in the required columns for each device:
		- device_id → The Device ID from the inventory list.
		- agent_id → The Agent ID managing the device from the inventory list.
		- driver_id → Leave blank for now (this will be filled in later).
		- username → The username needed for authentication on the device.
		- password → The corresponding password.

Step 3: Retrieve Custom Driver IDs
	Run the listCustomDrivers.ps1 script to generate a CSV file listing all available custom drivers in your Domotz account. This CSV will include:
		- Driver Name
		- Driver ID
	Use this file to locate the Driver IDs you need.

Step 4: Assign Custom Drivers in Devices.csv
	1. Copy the relevant driver_id values from the list of custom drivers into the driver_id column for each applicable device.
	2. Set the Sample Period in the sample_period column:
		Valid Options: 5m, 10m, 15m, 30m (default), 1hr, 2hr, 6hr, 12hr, 24hr
		- If left empty, the script will default to 30m (1800 seconds) unless the driver requires a different minimum sample period.
		- If the script detects that the sample period is below the driver’s minimum required sample period, it will automatically retry with the correct value.

Step 5: Apply Custom Drivers in Bulk
	Run the MassApplyScripts.ps1 script to apply the custom drivers to the devices defined in Devices.csv.
		- The script will log successful and failed associations.
		- If an error occurs due to a sample period conflict, the script will automatically retry with the correct minimum sample period.

Step 6: Review Logs
	A log file (Driver_Association_Log.txt) will be created in the same directory.
	The log file will contain:
		- All successful driver associations
		- Failures and their reasons (e.g., incorrect credentials, invalid driver, API issues)
		- Any automatic retries performed by the script
	Use this log file to troubleshoot any failed associations.

Troubleshooting
	- Script not running? Ensure PowerShell execution policy allows scripts (Set-ExecutionPolicy Unrestricted -Scope Process).
	- Devices not appearing in Devices.csv? Re-run Device-List-Script.ps1 and verify that the correct IDs were copied.
	- Drivers not appearing in the custom drivers CSV? Run listCustomDrivers.ps1 and ensure the correct account is used.
	- Failure in driver association? Check Driver_Association_Log.txt for details.

**Final Notes**
	- Always **review your `Devices.csv` carefully** before running the mass apply script.
	- The script **automatically handles sample period conflicts**, but you should still verify that all drivers were correctly applied.
	- **To associate multiple custom drivers with the same device, add a separate row for each driver in `Devices.csv`, using the same `device_id` and `agent_id` but specifying a different `driver_id` on each line.**
	- If you encounter persistent errors, refer to Driver_Association_Log.txt for debugging, or reach out to Bobby Wilson (bobby@domotz.com)
