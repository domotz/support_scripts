# DISCLAIMER:
# This script is provided "AS IS" and is intended solely for illustrative or educational purposes.
# Domotz makes no warranties, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or non-infringement. Use of this script is at your own risk.
# By using this script, you acknowledge and agree that Domotz shall not be liable for any direct, indirect, incidental, or consequential damages or losses arising from its use.
# You further agree to indemnify, defend, and hold harmless Domotz and its affiliates from and against any claims, liabilities, damages, or expenses resulting from your use or misuse of this script.
#
# In the event of any conflict between this disclaimer and any other agreement between you and Domotz, this disclaimer shall prevail with respect to the use of this script.

# ---------------------------------------------------------------------------------------------------------------------------

# Check if operation parameter is provided
param(
    [string]$operation
)

# Function to load environment variables from .env file
function Load-EnvFile {
    $envPath = Join-Path $PSScriptRoot ".env"
    if (Test-Path $envPath) {
        Get-Content $envPath | ForEach-Object {
            if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
                $key = $matches[1].Trim()
                # Remove any quotes from the value and trim whitespace
                $value = $matches[2].Trim().Trim('"', "'")
                Set-Item "env:$key" $value
            }
        }
        Write-Host "Environment variables loaded successfully from .env file"
    }
    else {
        Write-Host "WARNING: .env file not found at $envPath" -ForegroundColor Yellow
        Write-Host "Please create a .env file with the following content:"
        Write-Host "API_KEY=your_api_key_here"
        Write-Host "BASE_URL=your_base_url_here"
        exit
    }
}

# Load environment variables
Load-EnvFile

# Check if required environment variables are set
if (-not $env:API_KEY -or -not $env:BASE_URL) {
    Write-Host "ERROR: Required environment variables API_KEY and/or BASE_URL are not set in .env file" -ForegroundColor Red
    exit
}

# Define API Key and Base URL from environment variables and ensure proper formatting
$apiKey = $env:API_KEY
$baseURL = $env:BASE_URL.TrimEnd('/')  # Remove trailing slash if present

# Validate base URL format
try {
    $uri = [System.Uri]::new($baseURL)
    if (-not ($uri.Scheme -eq "http" -or $uri.Scheme -eq "https")) {
        throw "Invalid URL scheme. Must be http or https."
    }
}
catch {
    Write-Host "ERROR: Invalid BASE_URL format in .env file. Must be a valid HTTP/HTTPS URL." -ForegroundColor Red
    Write-Host "Current value: $baseURL"
    exit
}

# Function to show usage
function Show-Usage {
    $usageMessage = @"
USAGE: .\tcp_sensors_mass.ps1 -operation <operation_type>

OPERATION TYPES:
    list   : List all TCP sensors for devices in the CSV
    create : Create new TCP sensors for the specified ports
    delete : Delete TCP sensors for the specified ports

EXAMPLE:
    .\tcp_sensors_mass.ps1 -operation list
    .\tcp_sensors_mass.ps1 -operation create
    .\tcp_sensors_mass.ps1 -operation delete

NOTE: The script requires a 'tcp_sensors_mass_Devices.csv' file in the same directory with the following format:
    agent_id,device_ip,port_numbers_pipe_separated
"@
    Write-Host $usageMessage -ForegroundColor Yellow
    $usageMessage | Out-File -FilePath $logFile -Append
    exit
}

# Read CSV File
$csvPath = ".\tcp_sensors_mass_Devices.csv"
$logFile = ".\TCP_Sensors_Operation_Log.txt"

# Validate operation parameter
if ([string]::IsNullOrEmpty($operation)) {
    Show-Usage
}

# Validate operation value
$validOperations = @("list", "create", "delete")
if ($operation -notin $validOperations) {
    $errorMessage = "ERROR: Invalid operation '$operation'"
    Write-Host $errorMessage -ForegroundColor Red
    $errorMessage | Out-File -FilePath $logFile -Append
    Show-Usage
}


# Function to write log separator
function Write-LogHeader {
    $separator = "=" * 80
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    @"

$separator
NEW EXECUTION - $timestamp
Operation: $operation
$separator

"@ | Out-File -FilePath $logFile -Append
}

# Check if CSV file exists
if (-Not (Test-Path $csvPath)) {
    $errorMessage = @"
ERROR: CSV file 'tcp_sensors_mass_Devices.csv' not found in the current directory.
Please ensure that:
1. The file 'tcp_sensors_mass_Devices.csv' is in the same folder as this script
2. The file contains the following headers: agent_id,device_ip,port_numbers_pipe_separated
"@
    Write-Host $errorMessage -ForegroundColor Red
    $errorMessage | Out-File -FilePath $logFile -Append
    exit
}

# Initialize log file with header for this execution
Write-LogHeader

# Function to get device list for an agent
function Get-DeviceList {
    param (
        [string]$agentID
    )
    
    try {
        $apiEndpoint = "$baseURL/agent/$agentID/device"
        $headers = @{
            "Accept"    = "application/json"
            "X-Api-Key" = $apiKey
        }
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        return $response
    }
    catch {
        $errorMessage = "ERROR: Failed to get device list for Agent ID $agentID - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return $null
    }
}

# Function to get device ID from IP
function Get-DeviceIDFromIP {
    param (
        [string]$deviceIP,
        [string]$agentID,
        [array]$deviceList
    )
    
    $device = $deviceList | Where-Object { $_.ip_addresses -contains $deviceIP }
    if ($device) {
        $logMessage = "Mapped Device IP $deviceIP to Device ID $($device.id) on Agent ID $agentID"
        Write-Host $logMessage -NoNewline
        $logMessage | Out-File -FilePath $logFile -Append -NoNewline
        return $device.id
    }
    
    $errorMessage = "ERROR: No device found with IP $deviceIP on Agent ID $agentID"
    Write-Host "$errorMessage`n" -ForegroundColor Red
    "$errorMessage`n" | Out-File -FilePath $logFile -Append
    return $null
}

# Read CSV content and filter out comments and empty lines
$csvContent = Get-Content $csvPath | Where-Object { 
    $_ -match '\S' -and # Skip empty or whitespace-only lines
    -not $_.StartsWith('#')  # Skip comment lines
}

# Add header line back to content
$csvContent = @(
    "agent_id,device_ip,port_numbers_pipe_separated"  # Header
    $csvContent | Select-Object -Skip 1  # Skip original header
)

# Convert filtered content to CSV object
$devices = $csvContent | ConvertFrom-Csv

# Initialize counters
$script:totalAttempts = 0
$script:successCount = 0
$script:failureCount = 0
$script:successDetails = @()
$script:failureDetails = @()

# Cache for device lists to avoid repeated API calls
$deviceListCache = @{}

# Function to list TCP sensors
function list_tcp_ports_sensors {
    param (
        [string]$deviceID,
        [string]$agentID,
        [string]$deviceIP,
        [switch]$returnResponse
    )

    Write-Host "`nListing TCP sensors for Device IP $deviceIP (Device ID: $deviceID) on Agent ID $agentID..."
    "`nListing TCP sensors for Device IP $deviceIP (Device ID: $deviceID) on Agent ID $agentID..." | Out-File -FilePath $logFile -Append
    
    try {
        $apiEndpoint = "$baseURL/agent/$agentID/device/$deviceID/eye/tcp"
        $headers = @{
            "Accept"    = "application/json"
            "X-Api-Key" = $apiKey
        }

        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        
        if ($response.Count -gt 0) {
            Write-Host "TCP Sensors found:"
            "TCP Sensors found:" | Out-File -FilePath $logFile -Append
            
            # Format the table header
            $tableHeader = @"
Port Status Service ID Last Update
---- ------ ---------- -----------
"@
            Write-Host $tableHeader
            $tableHeader | Out-File -FilePath $logFile -Append
            
            # Format each sensor
            foreach ($sensor in $response) {
                $lastUpdate = if ($sensor.last_update) {
                    try { ([DateTime]($sensor.last_update.ToString())).ToString("yyyy-MM-dd HH:mm") }
                    catch { "N/A" }
                }
                else { "N/A" }
                
                $line = "{0,-4} {1,-6} {2,-10} {3,-11}" -f $sensor.port, $sensor.status, $sensor.id, $lastUpdate
                Write-Host $line
                $line | Out-File -FilePath $logFile -Append
            }
            Write-Host ""
            "" | Out-File -FilePath $logFile -Append
        }
        else {
            Write-Host "No TCP sensors found for this device`n"
            "No TCP sensors found for this device`n" | Out-File -FilePath $logFile -Append
        }
        
        if ($returnResponse) {
            return $response
        }
    }
    catch {
        $errorMessage = "ERROR: Failed to list TCP sensors for Device IP $deviceIP (Device ID: $deviceID) - $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        if ($returnResponse) {
            return $null
        }
    }
}

# Function to set TCP port sensors (create operation)
function set_tcp_ports_sensors {
    param (
        [string]$deviceID,
        [string]$agentID,
        [string]$deviceIP,
        [string]$portNumbers
    )

    $script:totalAttempts++
    $message = "`nAttempting to create TCP sensors for Device IP $deviceIP (Device ID: $deviceID) on Agent ID $agentID with ports: $portNumbers"
    Write-Host $message
    $message | Out-File -FilePath $logFile -Append

    $ports = $portNumbers -split '\|'
    foreach ($port in $ports) {
        if ([string]::IsNullOrWhiteSpace($port)) { continue }

        try {
            $apiEndpoint = "$baseURL/agent/$agentID/device/$deviceID/eye/tcp"
            $headers = @{
                "Accept"       = "application/json"
                "Content-Type" = "application/json"
                "X-Api-Key"    = $apiKey
            }
            $body = @{
                port = [int]$port
            } | ConvertTo-Json

            $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Post -Body $body
            $successMessage = "SUCCESS: Created TCP sensor for port $port on Device IP $deviceIP (Device ID: $deviceID)"
            Write-Host $successMessage -ForegroundColor Green
            $successMessage | Out-File -FilePath $logFile -Append
            $script:successCount++
            $script:successDetails += "Device IP: $deviceIP, Device ID: $deviceID, Agent ID: $agentID, Port: $port"
        }
        catch {
            $script:failureCount++
            $errorText = "Device IP: $deviceIP, Device ID: $deviceID, Agent ID: $agentID, Port: $port - Error: $_"
            Write-Host "FAILED: $errorText" -ForegroundColor Red
            "FAILED: $errorText" | Out-File -FilePath $logFile -Append
            $script:failureDetails += $errorText
        }
    }
    Write-Host ""
    "" | Out-File -FilePath $logFile -Append
}

# Function to delete TCP port sensors
function delete_tcp_ports_sensors {
    param (
        [string]$deviceID,
        [string]$agentID,
        [string]$deviceIP,
        [string]$portNumbers
    )

    Write-Host "`nProcessing deletion for Device IP $deviceIP (Device ID: $deviceID) on Agent ID $agentID..." -NoNewline
    "`nProcessing deletion for Device IP $deviceIP (Device ID: $deviceID) on Agent ID $agentID..." | Out-File -FilePath $logFile -Append -NoNewline
    
    $currentSensors = list_tcp_ports_sensors -deviceID $deviceID -agentID $agentID -deviceIP $deviceIP -returnResponse
    if ($null -eq $currentSensors) { return }

    $portsToDelete = $portNumbers -split '\|' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    foreach ($port in $portsToDelete) {
        $script:totalAttempts++
        $sensor = $currentSensors | Where-Object { $_.port -eq [int]$port }
        
        if ($sensor) {
            try {
                $apiEndpoint = "$baseURL/agent/$agentID/device/$deviceID/eye/tcp/$($sensor.id)"
                $headers = @{
                    "Accept"    = "application/json"
                    "X-Api-Key" = $apiKey
                }
                
                Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Delete
                $successMessage = "SUCCESS: Deleted TCP sensor for port $port (Service ID: $($sensor.id)) on Device IP $deviceIP (Device ID: $deviceID)"
                Write-Host $successMessage -ForegroundColor Green
                $successMessage | Out-File -FilePath $logFile -Append
                $script:successCount++
                $script:successDetails += "Device IP: $deviceIP, Device ID: $deviceID, Agent ID: $agentID, Port: $port, Service ID: $($sensor.id)"
            }
            catch {
                $script:failureCount++
                $errorText = "Device IP: $deviceIP, Device ID: $deviceID, Agent ID: $agentID, Port: $port - Error: $_"
                Write-Host "FAILED: $errorText" -ForegroundColor Red
                "FAILED: $errorText" | Out-File -FilePath $logFile -Append
                $script:failureDetails += $errorText
            }
        }
        else {
            $warningMessage = "WARNING: No TCP sensor found for port $port on Device IP $deviceIP (Device ID: $deviceID)"
            Write-Host $warningMessage -ForegroundColor Yellow
            $warningMessage | Out-File -FilePath $logFile -Append
        }
    }
    Write-Host ""
    "" | Out-File -FilePath $logFile -Append
}

# Main execution logic based on operation
switch ($operation) {
    "list" {
        $message = "`n=== Listing TCP Sensors for All Devices ===`n"
        Write-Host $message
        $message | Out-File -FilePath $logFile -Append
        foreach ($device in $devices) {
            # Get device list if not in cache
            if (-not $deviceListCache.ContainsKey($device.agent_id)) {
                $deviceListCache[$device.agent_id] = Get-DeviceList -agentID $device.agent_id
            }
            
            # Get device ID from IP
            $deviceID = Get-DeviceIDFromIP -deviceIP $device.device_ip -agentID $device.agent_id -deviceList $deviceListCache[$device.agent_id]
            if ($deviceID) {
                list_tcp_ports_sensors -deviceID $deviceID -agentID $device.agent_id -deviceIP $device.device_ip
            }
        }
    }
    "create" {
        $message = "`n=== Creating TCP Sensors for All Devices ===`n"
        Write-Host $message
        $message | Out-File -FilePath $logFile -Append
        foreach ($device in $devices) {
            # Get device list if not in cache
            if (-not $deviceListCache.ContainsKey($device.agent_id)) {
                $deviceListCache[$device.agent_id] = Get-DeviceList -agentID $device.agent_id
            }
            
            # Get device ID from IP
            $deviceID = Get-DeviceIDFromIP -deviceIP $device.device_ip -agentID $device.agent_id -deviceList $deviceListCache[$device.agent_id]
            if ($deviceID) {
                set_tcp_ports_sensors -deviceID $deviceID -agentID $device.agent_id -deviceIP $device.device_ip -portNumbers $device.port_numbers_pipe_separated
            }
        }
    }
    "delete" {
        $warningMessage = "`nWARNING: This will delete all TCP sensors specified in the CSV file. Are you sure you want to continue? (Y/N)"
        Write-Host $warningMessage
        $warningMessage | Out-File -FilePath $logFile -Append
        $confirmation = Read-Host
        if ($confirmation -eq "Y") {
            $message = "`n=== Deleting TCP Sensors for All Devices ===`n"
            Write-Host $message
            $message | Out-File -FilePath $logFile -Append
            foreach ($device in $devices) {
                # Get device list if not in cache
                if (-not $deviceListCache.ContainsKey($device.agent_id)) {
                    $deviceListCache[$device.agent_id] = Get-DeviceList -agentID $device.agent_id
                }
                
                # Get device ID from IP
                $deviceID = Get-DeviceIDFromIP -deviceIP $device.device_ip -agentID $device.agent_id -deviceList $deviceListCache[$device.agent_id]
                if ($deviceID) {
                    delete_tcp_ports_sensors -deviceID $deviceID -agentID $device.agent_id -deviceIP $device.device_ip -portNumbers $device.port_numbers_pipe_separated
                }
            }
        }
        else {
            $cancelMessage = "Operation cancelled by user"
            Write-Host $cancelMessage
            $cancelMessage | Out-File -FilePath $logFile -Append
            exit
        }
    }
}

# Summary based on operation
$summary = @"
--------------------------------------
OPERATION SUMMARY: $operation
Total Operations Attempted: $script:totalAttempts
"@

if ($operation -ne "list") {
    $summary += @"

Successful Operations: $script:successCount
Failed Operations: $script:failureCount

Successful Details:
$($script:successDetails | ForEach-Object { "- $_" } | Out-String)

Failed Details:
$($script:failureDetails | ForEach-Object { "- $_" } | Out-String)
"@
}

$summary += "--------------------------------------"

Write-Host $summary
$summary | Out-File -FilePath $logFile -Append

$logMessage = "`nLOG FILE: $logFile"
Write-Host $logMessage
$logMessage | Out-File -FilePath $logFile -Append