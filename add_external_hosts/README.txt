# Add External Hosts to Domotz Using PowerShell

## Overview
This PowerShell script reads a CSV file containing external hostnames/IP addresses and adds them as external hosts in **Domotz** using the **Domotz Public API**.

The script:
- Reads host data from a CSV file.
- Sends a **POST request** to the Domotz API to create the external hosts.
- Logs all API requests and responses in an `error_log.txt` file.
- Provides clear success/failure messages.

## Prerequisites
- A **Domotz API Key** with permission to manage external hosts.
- A **Domotz Agent ID** to associate the hosts with.
- PowerShell **v5.1 or later** installed on your system.
- A CSV file in the correct format (see below).

## Installation
1. **Download the script:**  
   Save the PowerShell script (`Add-Domotz-ExternalHosts.ps1`) in a directory of your choice.

2. **Prepare the CSV file:**  
   The script expects a CSV file with the following structure:

   ```csv
   Friendly Name,External Hostname/IP Address
   Google DNS,8.8.8.8
   Cloudflare,1.1.1.1
   My Website,www.example.com

Make sure the CSV file is saved in the same directory as the script.

Open PowerShell and navigate to the directory containing the script:
cd "C:\path\to\your\script"

Run the script:
.\Add-Domotz-ExternalHosts.ps1

Enter the required information when prompted:
Enter your Domotz API Key: (your-api-key-here)
Enter your Domotz Agent ID: (your-agent-id-here)
Enter the CSV file name (including .csv): hosts.csv

Review the output:
The script will display real-time updates on the API calls, including:

- The URL used for the API call.
- The request body.
- HTTP status codes.
- Success or failure messages.

Logging & Debugging
All actions and responses are logged in error_log.txt in the same directory.
If hosts are not being added:
Check if your API Key and Agent ID are correct.
Check if the hosts already exist in Domotz.
Open error_log.txt to see detailed responses from the API.

Expected API Responses
Status Code	Meaning	Description
201	Created	Host successfully added.
401	Unauthorized	API Key is missing or invalid.
403	Forbidden	API Key does not have permission.
500	Server Error	Domotz server issue; try again later.

Troubleshooting
No Hosts Are Added, but No Errors Appear
If the script completes but does not show "Successfully added" messages:
Check in Domotz UI if the hosts were actually added after refreshing.
The API does not return a response body, so status 201 Created is the only confirmation.

License
This script is provided as is with no warranties. Use it at your own risk.

