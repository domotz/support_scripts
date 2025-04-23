# PowerShell Script to Add External Hosts to Domotz with Proper Response Handling

# Prompt for user input
$apiKey = Read-Host "Enter your Domotz API Key"
$agentId = Read-Host "Enter your Domotz Agent ID"
$csvFile = Read-Host "Enter the CSV file name (including .csv)"

# Check if the CSV file exists
if (!(Test-Path $csvFile)) {
    Write-Host "Error: The specified CSV file does not exist." -ForegroundColor Red
    Pause
    exit
}

# Read CSV file
$csvData = Import-Csv -Path $csvFile

# Correct API Base URL
$baseURL = "https://api-us-east-1-cell-1.domotz.com/public-api/v1"

# Log file path
$logFile = "error_log.txt"

# Ensure the log file exists and clear it
New-Item -ItemType File -Path $logFile -Force | Out-Null
"" | Out-File -FilePath $logFile  # Clears previous logs

# Counters for success and failures
$successCount = 0
$failureCount = 0

# Iterate through each row in the CSV
foreach ($row in $csvData) {
    $friendlyName = $row."Friendly Name"
    $externalHost = $row."External Hostname/IP Address"

    if ([string]::IsNullOrWhiteSpace($friendlyName) -or [string]::IsNullOrWhiteSpace($externalHost)) {
        Write-Host "Skipping row with missing values: Friendly Name or External Host is empty." -ForegroundColor Yellow
        "Skipping row with missing values: Friendly Name or External Host is empty." | Out-File -Append -FilePath $logFile
        continue
    }

    # Prepare API request body
    $body = @"
{
    "host": "$externalHost",
    "name": "$friendlyName"
}
"@

    # Define headers
    $headers = @{
        "Content-Type" = "application/json"
        "X-Api-Key"    = $apiKey
    }

    # API Endpoint
    $url = "$baseURL/agent/$agentId/device/external-host"

    # Log request details
    Write-Host "-------------------------------------" -ForegroundColor Cyan
    Write-Host "Processing: $friendlyName ($externalHost)"
    Write-Host "URL: $url"
    Write-Host "Headers: $(ConvertTo-Json $headers -Depth 2)"
    Write-Host "Body: $body"
    Write-Host "-------------------------------------" -ForegroundColor Cyan

    @"
-------------------------------------
Processing: $friendlyName ($externalHost)
URL: $url
Headers: $(ConvertTo-Json $headers -Depth 2)
Body: $body
-------------------------------------
"@ | Out-File -Append -FilePath $logFile

    # Make API request and capture response
    try {
        $response = Invoke-WebRequest -Uri $url -Method Post -Headers $headers -Body $body -ContentType "application/json" -ErrorVariable httpError -ErrorAction SilentlyContinue

        # Capture HTTP Status Code
        $statusCode = $response.StatusCode

        # Log HTTP status
        Write-Host "HTTP Status: $statusCode" -ForegroundColor Yellow
        @"
HTTP Status: $statusCode
-------------------------------------
"@ | Out-File -Append -FilePath $logFile

        # Check if response is successful (201 Created)
        if ($statusCode -eq 201) {
            Write-Host "Successfully added: $friendlyName ($externalHost)" -ForegroundColor Green
            "Successfully added: $friendlyName ($externalHost)" | Out-File -Append -FilePath $logFile
            $successCount++
        } else {
            Write-Host "Unexpected response (Status: $statusCode). Check logs." -ForegroundColor Red
            $failureCount++
        }
    } catch {
        # Capture HTTP Error Details
        Write-Host "Request Failed! HTTP Error: $($_.Exception.Message)" -ForegroundColor Red
        "Request Failed! HTTP Error: $($_.Exception.Message)" | Out-File -Append -FilePath $logFile
        $failureCount++
    }
}

# Summary Report
Write-Host "-------------------------------------" -ForegroundColor Cyan
Write-Host "Script Execution Completed." -ForegroundColor Cyan
Write-Host "Successfully added: $successCount hosts" -ForegroundColor Green
Write-Host "Failed to add: $failureCount hosts (see error_log.txt for details)" -ForegroundColor Red
Write-Host "-------------------------------------" -ForegroundColor Cyan

@"
-------------------------------------
Script Execution Completed.
Successfully added: $successCount hosts
Failed to add: $failureCount hosts (see error_log.txt for details)
-------------------------------------
"@ | Out-File -Append -FilePath $logFile

# Keep window open to review output
Pause
