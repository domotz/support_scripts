# Define file paths
$csvFilePath = "device-snmp-updated.csv"
$logFilePath = "snmp_update_log.txt"

# Define API key and base URL
$apiKey = "YOUR-API-KEY"
$baseURL = "https://api-us-east-1-cell-1.domotz.com/public-api/v1"

# Clear the log file if it exists
if (Test-Path $logFilePath) { Clear-Content $logFilePath }

# Read CSV
$devices = Import-Csv -Path $csvFilePath

# Counters
$total = 0
$successCount = 0
$failureCount = 0

# Function to log messages
function Write-Log {
    param ([string]$message)
    Write-Host $message
    Add-Content -Path $logFilePath -Value $message
}

Write-Log "Starting SNMP authentication update process..."

foreach ($device in $devices) {
    $total++

    # Extract values from CSV with corrected column name
    $agentId = $device."Domotz Site/Agent ID"
    $deviceId = $device."Domotz Device ID"
    $version = $device."Version (V1, V2, V3)" -as [string]  # Ensure Version is a string
    $snmpRead = $device."SNMP Read"
    $snmpWrite = $device."SNMP Write"
    $username = $device.Username
    $authKey = $device."Authentication Key"
    $authProtocol = $device."Authentication Protocol (SHA, SHA-224, SHA-256, SHA-384, SHA-512)"
    $encKey = $device."Encryption Key"
    $encProtocol = $device."Encryption Protocol"

    # Trim and standardize SNMP version
    if ($version) {
        $version = $version.Trim().ToUpper()  # Ensure case matches API requirements
    }

    # Check if Version is missing
    if (-not $version) {
        Write-Log "[$total] ERROR: Missing SNMP version for Device ID: $deviceId."
        $failureCount++
        continue
    }

    # Construct API endpoint
    $url = "$baseURL/agent/$agentId/device/$deviceId/snmp-authentication"

    # Initialize empty body
    $body = @{}

    # Validation based on SNMP version
    if ($version -eq "V1" -or $version -eq "V2") {
        # Validate that SNMP Read and Write have values
        if (-not $snmpRead -or -not $snmpWrite) {
            Write-Log "[$total] ERROR: SNMP Read and SNMP Write must be provided for SNMP v1/v2 (Device ID: $deviceId)."
            $failureCount++
            continue
        }

        # Build request body
        $body = @{
            "snmp_read_community" = $snmpRead
            "snmp_write_community" = $snmpWrite
            "version" = $version
        }
        
        Write-Log "[$total] INFO: Processing SNMP v1/v2 for Device ID: $deviceId."
    
    } elseif ($version -eq "V3") {
        # Validate that all required fields for SNMP v3 are provided
        $missingFields = @()
        if (-not $username) { $missingFields += "Username" }
        if (-not $authProtocol) { $missingFields += "Authentication Protocol" }
        if (-not $authKey) { $missingFields += "Authentication Key" }
        if (-not $encProtocol) { $missingFields += "Encryption Protocol" }
        if (-not $encKey) { $missingFields += "Encryption Key" }

        if ($missingFields.Count -gt 0) {
            Write-Log "[$total] ERROR: Missing required fields for SNMP v3 (Device ID: $deviceId)."
            Write-Host "Device $deviceId requires the following missing fields: $($missingFields -join ', ')"

            # Prompt user for missing values
            foreach ($field in $missingFields) {
                $value = Read-Host "Enter value for $field"
                switch ($field) {
                    "Username" { $username = $value }
                    "Authentication Protocol" { $authProtocol = $value }
                    "Authentication Key" { $authKey = $value }
                    "Encryption Protocol" { $encProtocol = $value }
                    "Encryption Key" { $encKey = $value }
                }
            }
        }

        # Build request body for SNMP v3
        $body = @{
            "username" = $username
            "authentication_key" = $authKey
            "authentication_protocol" = $authProtocol
            "encryption_key" = $encKey
            "encryption_protocol" = $encProtocol
            "version" = "V3_AUTH_PRIV"
        }
        
        Write-Log "[$total] INFO: Processing SNMP v3 for Device ID: $deviceId."
    
    } else {
        Write-Log "[$total] ERROR: Unsupported SNMP version ($version) for Device ID: $deviceId."
        $failureCount++
        continue
    }

    # Convert body to JSON
    $jsonBody = $body | ConvertTo-Json -Depth 2

    # Perform API request
    try {
        $response = Invoke-RestMethod -Uri $url -Method Put -Headers @{
            "Content-Type" = "application/json"
            "X-Api-Key" = $apiKey
        } -Body $jsonBody

        Write-Log "[$total] SUCCESS: SNMP updated for Device ID: $deviceId (Agent $agentId)."
        $successCount++
    } catch {
        Write-Log "[$total] FAILURE: Error updating SNMP for Device ID: $deviceId (Agent $agentId). Error: $_"
        $failureCount++
    }
}

# Summary
Write-Log "-----------------------------------"
Write-Log "Total Attempts: $total"
Write-Log "Successful Updates: $successCount"
Write-Log "Failed Updates: $failureCount"
Write-Log "-----------------------------------"
Write-Log "Process completed. Logs saved in $logFilePath."

# Prevent window from closing
Read-Host "Press Enter to exit"
