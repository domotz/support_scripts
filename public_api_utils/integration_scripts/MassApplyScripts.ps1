# Define API Key and Base URL
$apiKey = "YOUR-API-KEY"
$baseURL = "https://api-us-east-1-cell-1.domotz.com/public-api/v1"

# Read CSV File (Ensure it's in the same directory)
$csvPath = ".\Devices.csv"
$logFile = ".\Driver_Association_Log.txt"

# Sample period conversion mapping
$samplePeriodMap = @{
    "5m"   = 300
    "10m"  = 600
    "15m"  = 900
    "30m"  = 1800
    "1hr"  = 3600
    "2hr"  = 7200
    "6hr"  = 21600
    "12hr" = 43200
    "24hr" = 86400
}

# Initialize log file
"Domotz Custom Driver Association Log - $(Get-Date)" | Out-File -FilePath $logFile

if (-Not (Test-Path $csvPath)) {
    Write-Host "ERROR: CSV file '$csvPath' not found. Exiting."
    "Error: CSV file '$csvPath' not found." | Out-File -FilePath $logFile -Append
    Read-Host "Press Enter to exit"
    exit
}

$devices = Import-Csv -Path $csvPath

# Check if CSV contains required columns
$requiredColumns = @("device_id", "agent_id", "driver_id", "username", "password")
$missingColumns = $requiredColumns | Where-Object {$_ -notin $devices[0].PSObject.Properties.Name}
if ($missingColumns.Count -gt 0) {
    Write-Host "ERROR: Missing required columns in CSV: $($missingColumns -join ', ')"
    "Error: Missing required columns in CSV: $($missingColumns -join ', ')" | Out-File -FilePath $logFile -Append
    Read-Host "Press Enter to exit"
    exit
}

# Counters for Summary (Now using script-wide variables)
$script:totalAttempts = 0
$script:successCount = 0
$script:failureCount = 0
$script:alreadyAssociatedCount = 0
$failureDetails = @()

# Function to associate a driver with retry logic
function Associate-Driver {
    param (
        [string]$deviceID,
        [string]$agentID,
        [string]$driverID,
        [string]$username,
        [string]$password,
        [int]$samplePeriod,
        [bool]$retrying
    )

    $script:totalAttempts++
    $retryText = if ($retrying) {"(Retry) "} else {""}
    Write-Host "$retryText Attempting to associate Driver ID $driverID to Device ID $deviceID on Agent ID $agentID with sample period $samplePeriod seconds..."
    "$retryText Attempting: Driver ID $driverID -> Device ID $deviceID (Agent ID $agentID) | Sample Period: $samplePeriod" | Out-File -FilePath $logFile -Append

    # Set API Endpoint for Association
    $apiEndpoint = "$baseURL/custom-driver/$driverID/agent/$agentID/device/$deviceID/association"

    # Define Headers
    $headers = @{
        "Accept" = "application/json"
        "Content-Type" = "application/json"
        "X-Api-Key" = $apiKey
    }

    # Define API Request Body
    $body = @{
        credentials = @{
            username = $username
            password = $password
        }
        parameters = @()
        sample_period = $samplePeriod
    } | ConvertTo-Json -Depth 10

    # Perform API Request
    try {
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Post -Body $body
        Write-Host "$retryText SUCCESS: Associated Driver ID $driverID to Device ID $deviceID on Agent ID $agentID"
        "$retryText Success: Driver ID $driverID -> Device ID $deviceID (Agent ID $agentID)" | Out-File -FilePath $logFile -Append
        $script:successCount++
    } catch {
        # Extract error message
        $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue

        if ($errorMessage) {
            # Handle "Driver already associated" scenario
            if ($errorMessage.error -match "Driver $driverID and device $deviceID are already associated") {
                Write-Host "INFO: Driver ID $driverID is already associated with Device ID $deviceID."
                "INFO: Driver ID $driverID is already associated with Device ID $deviceID." | Out-File -FilePath $logFile -Append
                $script:alreadyAssociatedCount++
                return
            }
            # Handle minimum sample period requirement
            elseif ($errorMessage.error -match "sample_period must be equal to or greater than the minimal_sample_period of the driver") {
                if ($errorMessage.error -match "\((\d+)\)") {
                    $minimalSamplePeriod = [int]$matches[1]
                    if (-not $retrying) {
                        Write-Host "Retrying with minimal sample period: $minimalSamplePeriod seconds..."
                        Associate-Driver -deviceID $deviceID -agentID $agentID -driverID $driverID -username $username -password $password -samplePeriod $minimalSamplePeriod -retrying $true
                        return
                    }
                }
            }
        }
        
        # If failure occurs, count it
        if (-not $retrying) {
            $script:failureCount++
        }

        $errorText = "FAILED: Driver ID $driverID -> Device ID $deviceID on Agent ID $agentID | Error: $_"
        Write-Host $errorText
        $failureDetails += $errorText
    }
}

# Loop through each row and apply the Custom Driver
foreach ($device in $devices) {
    $deviceID = $device.device_id
    $agentID = $device.agent_id
    $driverID = $device.driver_id
    $username = $device.username
    $password = $device.password

    # Determine sample period (default to 1800 if missing or empty)
    $samplePeriodStr = $device.sample_period
    if (-not $samplePeriodStr -or -not $samplePeriodMap.ContainsKey($samplePeriodStr)) {
        $samplePeriod = 1800
    } else {
        $samplePeriod = $samplePeriodMap[$samplePeriodStr]
    }

    # Associate driver (with retry handling inside function)
    Associate-Driver -deviceID $deviceID -agentID $agentID -driverID $driverID -username $username -password $password -samplePeriod $samplePeriod -retrying $false
}

# Summary
$summary = @"
--------------------------------------
PROCESS COMPLETED
Total Attempts: $script:totalAttempts
Successes: $script:successCount
Failures: $script:failureCount
Already Associated: $script:alreadyAssociatedCount
--------------------------------------
"@

Write-Host $summary
$summary | Out-File -FilePath $logFile -Append

# Log failure details if any
if ($failureDetails.Count -gt 0) {
    "`nFailure Details:`n" | Out-File -FilePath $logFile -Append
    $failureDetails | Out-File -FilePath $logFile -Append
}

Write-Host "SUMMARY WRITTEN TO: $logFile"
Read-Host "Press Enter to exit"