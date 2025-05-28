# TCP Sensors Mass Management Script

This script allows you to manage TCP sensors for multiple devices in bulk using the Domotz Public API. You can list, create, and delete TCP sensors for multiple devices specified in a CSV file.

## Prerequisites

1. PowerShell 5.1 or higher
2. A valid Domotz API Key
3. Ensure PowerShell Execution Policy allows script execution.
   If needed, run the following in PowerShell (as Administrator):
   ```powershell
   Set-ExecutionPolicy Unrestricted -Scope Process
   ```

## Setup

### 1. Environment File (.env)

Create a `.env` file in the same directory as the script with the following content:

```
API_KEY='your_api_key_here'
BASE_URL='https://api-us-east-1-cell-1.domotz.com/public-api/v1'
```

Important notes for the .env file:

- Replace `your_api_key_here` with your actual Domotz API key
- Do NOT use quotes around the values
- Do NOT add spaces before or after the equals sign
- Values should be specified exactly as shown above

Example of correct format:

```
API_KEY='OBGmiJvxxxxxxxxxxk0dX2sNKKpmkYgRdfc2zBKgU'
BASE_URL='https://api-us-east-1-cell-1.domotz.com/public-api/v1'
```

Example of incorrect format (DO NOT USE):

```
API_KEY="OBGmiJvmilBMxxxxxxxsNKKpmkYgRdfc2zBKgU"
BASE_URL = "https://api-us-east-1-cell-1.domotz.com/public-api/v1"
```

### 2. CSV File (tcp_sensors_mass_Devices.csv)

Create a CSV file named `tcp_sensors_mass_Devices.csv` with the following format:

```csv
agent_id,device_ip,port_numbers_pipe_separated
311833,10.10.200.245,80|443|8080
311833,10.10.221.1,22|3389
```

- `agent_id`: The Domotz Agent ID where the device is located
- `device_ip`: The IP address of the device
- `port_numbers_pipe_separated`: TCP ports to monitor, separated by pipe character (|)

This .csv accepet # as comments.

## Usage

The script supports three operations:

1. **List TCP Sensors**

   ```powershell
   .\tcp_sensors_mass.ps1 -operation list
   ```

   Lists all TCP sensors for devices specified in the CSV file.

2. **Create TCP Sensors**

   ```powershell
   .\tcp_sensors_mass.ps1 -operation create
   ```

   Creates new TCP sensors for the specified ports in the CSV file.

3. **Delete TCP Sensors**
   ```powershell
   .\tcp_sensors_mass.ps1 -operation delete
   ```
   Deletes TCP sensors for the specified ports in the CSV file.

## Log Files

The script generates log files to track operations:

- `TCP_Sensors_Operation_Log.txt`: General operation log

## Error Handling

- The script validates the existence of required files (.env and CSV)
- Invalid or unreachable devices are logged with error messages
- API errors are captured and logged
- Operation summaries are provided at the end of each run

## Notes

- Ensure your API key has the necessary permissions
- The script will skip empty lines and comments in the CSV file
- For the delete operation, you'll be prompted for confirmation before proceeding
- Make sure the device IPs in the CSV file are reachable from the Domotz agent

## Support

For any issues or questions:

1. Check the error messages in the log files
2. Verify your API key and permissions
3. Ensure the devices are reachable from the Domotz agent
