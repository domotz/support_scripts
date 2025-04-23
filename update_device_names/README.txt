# Domotz Device Management Scripts

This repository contains two scripts to manage devices monitored by Domotz:
1. `list_devices_with_different_names.py`: Lists devices where the Name is different from the DHCP name.
2. `update_device_names_from_csv.py`: Updates device names to match the DHCP names based on a provided CSV file.

## Prerequisites

- Python 3.x
- The `requests` library (`pip install requests`)

## Script 1: list_devices_with_different_names.py

### Description

This script fetches all devices monitored by Domotz agents and lists those where the Name is different from the DHCP name. The output is saved to a CSV file named `devices_with_different_names.csv`.

### Fields to Update

- **API_KEY**: Replace `'YOUR_API_KEY'` with your actual Domotz API key.
- **API_URL**: Update the API endpoint URL based on your location:
  - North American users: `'https://api-us-east-1-cell-1.domotz.com/public-api/v1'`
  - Users outside of North America: `'https://api-eu-west-1-cell-1.domotz.com/public-api/v1'`

### Steps to Run

1. Open the script file `list_devices_with_different_names.py`.
2. Update the `API_KEY` and `API_URL` fields as needed.
3. Open a command prompt and navigate to the directory containing the script.
4. Run the script with `python list_devices_with_different_names.py`.
