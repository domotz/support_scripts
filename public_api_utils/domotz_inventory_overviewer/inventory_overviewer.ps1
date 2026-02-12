# DISCLAIMER:
# This script is provided "AS IS" and is intended solely for illustrative or educational purposes.
# Domotz makes no warranties, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or non-infringement. Use of this script is at your own risk.
# By using this script, you acknowledge and agree that Domotz shall not be liable for any direct, indirect, incidental, or consequential damages or losses arising from its use.
# You further agree to indemnify, defend, and hold harmless Domotz and its affiliates from and against any claims, liabilities, damages, or expenses resulting from your use or misuse of this script.
#
# In the event of any conflict between this disclaimer and any other agreement between you and Domotz, this disclaimer shall prevail with respect to the use of this script.

# ---------------------------------------------------------------------------------------------------------------------------
#
# SCRIPT PURPOSE:
# This script extracts data from Domotz (Organizations, Collectors, Devices) and exports to Excel.
#
# PARAMETERS:
#   -object-type   : What to extract: "all", "organizations", "collectors", "devices" (comma-separated)
#   -collector_ids : Comma-separated list of collector IDs to filter (optional, default: all)
#   -operation     : "extract" or "list_collectors" (default: extract)
#   -device-type   : "managed", "unmanaged", or "managed,unmanaged" (default: managed,unmanaged)
#
# ---------------------------------------------------------------------------------------------------------------------------

# ================================================================================
# # SETUP INSTRUCTIONS (First Time)
# ================================================================================

# 1. CONFIGURE .ENV FILE:
# Create a file named '.env' in the same folder as this script with:

# API_KEY='<your API Key here>'
# BASE_URL='<API endpoint for your region>'

# BASE_URL Options:
# - US Region: https://api-us-east-1-cell-1.domotz.com/public-api/v1
# - EU Region: https://api-eu-west-1-cell-1.domotz.com/public-api/v1

# 2. CHECK IMPORTEXCEL MODULE:
# Run: Get-Module -ListAvailable -Name ImportExcel

# If the command returns nothing, the module is NOT installed.
# Install it with: Install-Module -Name ImportExcel -Scope CurrentUser -Force


# Check if operation parameter is provided
param(
    [string]$collector_ids,
    [string]$operation,
    [Alias("device-type")]
    [string]$device_type,
    [string]$filename,
    [switch]$debug,
    [Alias("h", "?")]
    [switch]$help
)

# Check for help arguments (support both / and - prefixes)
if ($args -contains "/help" -or $args -contains "/?" -or $args -contains "/h") {
    # Set a flag to show help after functions are defined
    $script:showHelpOnly = $true
}

# Check for unknown/invalid parameters passed via $args
if ($args.Count -gt 0 -and -not $script:showHelpOnly) {
    # Filter out help flags that we already handled
    $unknownArgs = $args | Where-Object { $_ -notin @("/help", "/?", "/h") }
    if ($unknownArgs.Count -gt 0) {
        $script:hasInvalidParams = $true
        $script:invalidParamsList = $unknownArgs
    }
}

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

# Check if ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    $errorMsg = @"

================================================================================
                                    ERROR                                       
================================================================================

The ImportExcel PowerShell module is required for this script!

SOLUTION: Install the module by running the following command:

Install-Module -Name ImportExcel -Scope CurrentUser

Script execution stopped.
================================================================================
"@
    Write-Host $errorMsg -ForegroundColor Red
    exit
}

# Get the PowerShell script name dynamically (without .ps1 extension)
$PS_SCRIPT_NAME = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

# Define constants
$DEFAULT_EXCEL_BASENAME = $PS_SCRIPT_NAME
$DEFAULT_EXCEL_FILENAME = "$DEFAULT_EXCEL_BASENAME.xlsx"

# Define log file (dynamic based on script name)
$logFile = ".\$($PS_SCRIPT_NAME)_Log.txt"

# Define pagination constant for agent list retrieval
$AGENT_PAGE_SIZE = 100

# Helper function to retrieve all agents with pagination
function Get-AllAgents {
    $allAgents = @()
    $pageNumber = 0
    
    do {
        try {
            $agentEndpoint = "$baseURL/agent?page_size=$AGENT_PAGE_SIZE&page_number=$pageNumber"
            $agentHeaders = @{
                "X-Api-Key"    = $apiKey
                "Accept"       = "application/json"
                "Content-Type" = "application/json"
            }
            
            if ($debug) {
                $debugRequestMsg = "`n[DEBUG] Requesting agents - Page $pageNumber"
                $debugRequestMsg += "`n[DEBUG] GET $agentEndpoint"
                Write-Host $debugRequestMsg -ForegroundColor Cyan
                $debugRequestMsg | Out-File -FilePath $logFile -Append
            }
            
            $agents = Invoke-RestMethod -Uri $agentEndpoint -Method Get -Headers $agentHeaders
            
            if ($debug) {
                $debugResponseMsg = "[DEBUG] Response - Received $($agents.Count) agent(s) on page $pageNumber"
                if ($agents.Count -gt 0) {
                    $debugResponseMsg += "`n[DEBUG] Response Data: $($agents | ConvertTo-Json -Compress -Depth 2)"
                }
                Write-Host $debugResponseMsg -ForegroundColor Cyan
                $debugResponseMsg | Out-File -FilePath $logFile -Append
            }
            
            if ($agents.Count -eq 0) {
                break
            }
            
            $allAgents += $agents
            $pageNumber++
        }
        catch {
            $errorMsg = "`nERROR: Failed to retrieve agents on page $pageNumber - $_"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -FilePath $logFile -Append
            break
        }
    } while ($agents.Count -gt 0)
    
    return $allAgents
}

# Function to list collectors/agents
function List-Collectors {
    param (
        [bool]$numbered = $false,
        [bool]$silent = $false
    )
    
    # Use pagination helper to get all agents
    $agents = Get-AllAgents
    
    if ($agents.Count -eq 0) {
        if (-not $silent) {
            $noAgentsMsg = "`nNo collectors/agents found in your Domotz account."
            Write-Host $noAgentsMsg -ForegroundColor Yellow
            $noAgentsMsg | Out-File -FilePath $logFile -Append
        }
        return @()
    }
    else {
        $sortedAgents = $agents | Sort-Object display_name
        
        if (-not $silent) {
            $agentHeaderMsg = @"

================================================================================
AVAILABLE COLLECTORS/AGENTS
================================================================================
"@
            Write-Host $agentHeaderMsg -ForegroundColor Green
            $agentHeaderMsg | Out-File -FilePath $logFile -Append
            
            $index = 1
            
            foreach ($agent in $sortedAgents) {
                if ($numbered) {
                    $agentLine = "  [$index] '$($agent.display_name)' (ID: $($agent.id))"
                }
                else {
                    $agentLine = "  - '$($agent.display_name)' (ID: $($agent.id))"
                }
                Write-Host $agentLine
                $agentLine | Out-File -FilePath $logFile -Append
                $index++
            }
            
            $agentSummaryMsg = "`nTotal: $($agents.Count) collector(s)/agent(s) found."
            Write-Host $agentSummaryMsg -ForegroundColor Yellow
            $agentSummaryMsg | Out-File -FilePath $logFile -Append
        }
        
        return $sortedAgents
    }
}

# Function to show help (usage only, no interactive workflow)
function Show-Help {
    $usageMessage = @"
================================================================================
        DOMOTZ INVENTORY OVERVIEWER - EXPORT TO EXCEL TOOL
================================================================================

USAGE: .\$PS_SCRIPT_NAME.ps1 [-operation <operation_type>] [additional parameters]

================================================================================
OPERATION TYPES
================================================================================

    list_collectors : List all available collectors/agents
                      No additional parameters required

    extract         : Extract data from Domotz and export to Excel (default)
                      Optional: -collector_ids <ids> -device-type <types>

================================================================================
PARAMETERS
================================================================================

    -operation      : Operation to perform
                      Values: "extract" (default), "list_collectors"

    -collector_ids  : Filter by specific collector IDs (comma-separated)
                      If not specified, all collectors are included
                      NOTE: When specified, Organizations worksheet is not created
                      Example: -collector_ids "312189,313759"

    -device-type    : Filter device types (comma-separated)
                      Values: "managed", "unmanaged", "managed,unmanaged" (default)

    -filename       : Custom output Excel filename (optional)
                      Example: -filename "my_export"

    -debug          : Enable detailed logging

================================================================================
EXAMPLES
================================================================================

List all collectors:
.\$PS_SCRIPT_NAME.ps1 -operation list_collectors

Extract all data (organizations, collectors, devices):
.\$PS_SCRIPT_NAME.ps1

Extract data from specific collectors (no Organizations worksheet):
.\$PS_SCRIPT_NAME.ps1 -collector_ids "312189,313759"

Extract only managed devices:
.\$PS_SCRIPT_NAME.ps1 -device-type managed

Extract with custom filename:
.\$PS_SCRIPT_NAME.ps1 -filename "domotz_export_2025"

================================================================================
OUTPUT
================================================================================

The script creates an Excel file with the following worksheets:

- Organizations : List of organizations (only when -collector_ids is NOT specified)
- Collectors    : List of collectors/agents
- Devices       : List of devices with device-type column
                  The device-type column indicates "managed" or "unmanaged"

================================================================================
"@
    Write-Host $usageMessage -ForegroundColor Yellow
    exit
}

# Function to show usage
function Show-Usage {
    $usageMessage = @"
================================================================================
        DOMOTZ INVENTORY OVERVIEWER - EXPORT TO EXCEL TOOL
================================================================================

USAGE: .\$PS_SCRIPT_NAME.ps1 [-operation <operation_type>] [additional parameters]

================================================================================
QUICK START
================================================================================

Extract all data (default):
.\$PS_SCRIPT_NAME.ps1

List available collectors:
.\$PS_SCRIPT_NAME.ps1 -operation list_collectors

For full help:
.\$PS_SCRIPT_NAME.ps1 -help

================================================================================
"@
    Write-Host $usageMessage -ForegroundColor Yellow
    $usageMessage | Out-File -FilePath $logFile -Append
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

# Function to open Excel file
function Open-Excel {
    param (
        [string]$fileName
    )
    
    $message = "`n=== Opening Excel File ===`n"
    Write-Host $message -ForegroundColor Magenta
    $message | Out-File -FilePath $logFile -Append
    
    # Determine which file to open
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = $DEFAULT_EXCEL_FILENAME
    }
    
    # Add .xlsx extension if not present
    if (-not $fileName.EndsWith(".xlsx")) {
        $fileName = "$fileName.xlsx"
    }
    
    $excelPath = Join-Path $PSScriptRoot $fileName
    
    if (-not (Test-Path $excelPath)) {
        $errorMsg = "ERROR: File not found: $fileName"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw "File not found: $excelPath"
    }
    
    try {
        $openMsg = "Opening file: $fileName"
        Write-Host $openMsg -ForegroundColor Cyan
        $openMsg | Out-File -FilePath $logFile -Append
        
        # Open the file with the default application
        Start-Process $excelPath
        
        $successMsg = "[OK] Excel file opened successfully"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
    }
    catch {
        $errorMsg = "ERROR: Failed to open Excel file - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw $_
    }
}

# Function to get managed devices for a collector
function Get-ManagedDevices {
    param (
        [string]$collectorID,
        [string]$collectorName
    )
    
    try {
        $apiEndpoint = "$baseURL/agent/$collectorID/device"
        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
            "X-Api-Key"    = $apiKey
        }
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        
        # Add collector info and device-type to each device
        $devices = @()
        foreach ($device in $response) {
            $device | Add-Member -NotePropertyName "_collector_id_" -NotePropertyValue $collectorID -Force
            $device | Add-Member -NotePropertyName "_collector_name_" -NotePropertyValue $collectorName -Force
            $device | Add-Member -NotePropertyName "_device_type_" -NotePropertyValue "managed" -Force
            $devices += $device
        }
        
        return $devices
    }
    catch {
        $errorMessage = "ERROR: Failed to get managed devices for Collector ID $collectorID - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return @()
    }
}

# Function to get unmanaged devices for a collector
function Get-UnmanagedDevices {
    param (
        [string]$collectorID,
        [string]$collectorName
    )
    
    try {
        $apiEndpoint = "$baseURL/agent/$collectorID/device/monitoring-state/unmanaged"
        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
            "X-Api-Key"    = $apiKey
        }
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        
        # Add collector info and device-type to each device
        $devices = @()
        foreach ($device in $response) {
            $device | Add-Member -NotePropertyName "_collector_id_" -NotePropertyValue $collectorID -Force
            $device | Add-Member -NotePropertyName "_collector_name_" -NotePropertyValue $collectorName -Force
            $device | Add-Member -NotePropertyName "_device_type_" -NotePropertyValue "unmanaged" -Force
            $devices += $device
        }
        
        return $devices
    }
    catch {
        $errorMessage = "ERROR: Failed to get unmanaged devices for Collector ID $collectorID - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return @()
    }
}

# Function to extract data and create Excel
function Extract-Data {
    param (
        [string]$collectorIds,
        [string]$deviceType,
        [string]$fileName
    )
    
    $message = "`n=== Extracting Domotz Data ===`n"
    Write-Host $message -ForegroundColor Magenta
    $message | Out-File -FilePath $logFile -Append
    
    # Always extract all object types, but Organizations only when collector_ids is NOT specified
    $hasCollectorFilter = -not [string]::IsNullOrEmpty($collectorIds)
    
    if ($hasCollectorFilter) {
        $objectTypes = @("collectors", "devices")
        $filterNoteMsg = "[INFO] Collector filter specified - Organizations worksheet will be skipped"
        Write-Host $filterNoteMsg -ForegroundColor Yellow
        $filterNoteMsg | Out-File -FilePath $logFile -Append
    }
    else {
        $objectTypes = @("organizations", "collectors", "devices")
    }
    
    $objTypesMsg = "[INFO] Object types to extract: $($objectTypes -join ', ')"
    Write-Host $objTypesMsg -ForegroundColor Cyan
    $objTypesMsg | Out-File -FilePath $logFile -Append
    
    # Parse device-type parameter
    if ([string]::IsNullOrEmpty($deviceType)) {
        $deviceTypes = @("managed", "unmanaged")
    }
    else {
        $deviceTypes = $deviceType.ToLower() -split ',' | ForEach-Object { $_.Trim() }
    }
    
    if ($objectTypes -contains "devices") {
        $devTypesMsg = "[INFO] Device types to extract: $($deviceTypes -join ', ')"
        Write-Host $devTypesMsg -ForegroundColor Cyan
        $devTypesMsg | Out-File -FilePath $logFile -Append
    }
    
    # STEP 1: Get all collectors from API (with pagination)
    $step1Msg = "`n[STEP 1] Retrieving collectors from Domotz API..."
    Write-Host $step1Msg -ForegroundColor Cyan
    $step1Msg | Out-File -FilePath $logFile -Append
    
    $allCollectors = Get-AllAgents
    
    if ($allCollectors.Count -eq 0) {
        $errorMsg = "ERROR: No collectors found or failed to retrieve collectors from API"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        return
    }
    
    $foundMsg = "[OK] Retrieved $($allCollectors.Count) collector(s) from API"
    Write-Host $foundMsg -ForegroundColor Green
    $foundMsg | Out-File -FilePath $logFile -Append
    
    # STEP 2: Filter collectors if collector_ids parameter is provided
    $step2Msg = "`n[STEP 2] Filtering collectors..."
    Write-Host $step2Msg -ForegroundColor Cyan
    $step2Msg | Out-File -FilePath $logFile -Append
    
    if (-not [string]::IsNullOrEmpty($collectorIds)) {
        $collectorIdArray = $collectorIds -split ',' | ForEach-Object { $_.Trim() }
        $filteredCollectors = $allCollectors | Where-Object { $collectorIdArray -contains $_.id.ToString() }
        
        if ($filteredCollectors.Count -eq 0) {
            $errorMsg = "ERROR: None of the specified collector IDs were found: $collectorIds"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -FilePath $logFile -Append
            
            $availableMsg = "Available collector IDs:"
            Write-Host $availableMsg -ForegroundColor Yellow
            $availableMsg | Out-File -FilePath $logFile -Append
            
            foreach ($coll in $allCollectors) {
                $collMsg = "  - $($coll.id): $($coll.display_name)"
                Write-Host $collMsg
                $collMsg | Out-File -FilePath $logFile -Append
            }
            return
        }
        
        $filterMsg = "[OK] Filtered to $($filteredCollectors.Count) collector(s) based on -collector_ids parameter"
        Write-Host $filterMsg -ForegroundColor Green
        $filterMsg | Out-File -FilePath $logFile -Append
        
        $collectors = $filteredCollectors
    }
    else {
        $noFilterMsg = "[INFO] No collector filter applied - using all $($allCollectors.Count) collector(s)"
        Write-Host $noFilterMsg -ForegroundColor Cyan
        $noFilterMsg | Out-File -FilePath $logFile -Append
        $collectors = $allCollectors
    }
    
    # Display collectors being processed
    $collListMsg = "Collectors to process:"
    Write-Host $collListMsg -ForegroundColor Yellow
    $collListMsg | Out-File -FilePath $logFile -Append
    
    foreach ($coll in $collectors) {
        $collLine = "  - ID: $($coll.id) - '$($coll.display_name)'"
        Write-Host $collLine
        $collLine | Out-File -FilePath $logFile -Append
    }
    
    # STEP 3: Prepare data collections
    $organizationsData = @()
    $collectorsData = @()
    $devicesData = @()
    $totalManagedDevices = 0
    $totalUnmanagedDevices = 0
    
    # STEP 3.1: Fetch managed device counts for all collectors (needed for Organizations and Collectors worksheets)
    $step3aMsg = "`n[STEP 3] Fetching managed device counts for all collectors..."
    Write-Host $step3aMsg -ForegroundColor Cyan
    $step3aMsg | Out-File -FilePath $logFile -Append
    
    $managedDeviceCountByCollector = @{}
    foreach ($coll in $collectors) {
        $countMsg = "  Counting managed devices for collector: $($coll.display_name) (ID: $($coll.id))"
        Write-Host $countMsg -ForegroundColor Gray
        $countMsg | Out-File -FilePath $logFile -Append
        
        try {
            $apiEndpoint = "$baseURL/agent/$($coll.id)/device"
            $headers = @{
                "Accept"       = "application/json"
                "Content-Type" = "application/json"
                "X-Api-Key"    = $apiKey
            }
            
            $managedDevicesResponse = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
            $managedDeviceCountByCollector[$coll.id] = $managedDevicesResponse.Count
            
            $countResultMsg = "    [OK] Found $($managedDevicesResponse.Count) managed device(s)"
            Write-Host $countResultMsg -ForegroundColor Green
            $countResultMsg | Out-File -FilePath $logFile -Append
        }
        catch {
            $managedDeviceCountByCollector[$coll.id] = 0
            $countErrorMsg = "    [WARNING] Failed to count managed devices: $_"
            Write-Host $countErrorMsg -ForegroundColor Yellow
            $countErrorMsg | Out-File -FilePath $logFile -Append
        }
    }
    
    # STEP 4: Extract Organizations data
    if ($objectTypes -contains "organizations") {
        $step4Msg = "`n[STEP 4] Extracting Organizations data..."
        Write-Host $step4Msg -ForegroundColor Cyan
        $step4Msg | Out-File -FilePath $logFile -Append
        
        # Extract unique organizations from collectors
        $orgMap = @{}
        foreach ($coll in $collectors) {
            if ($coll.organization -and $coll.organization.id) {
                $orgId = $coll.organization.id
                $collectorManagedCount = $managedDeviceCountByCollector[$coll.id]
                
                if (-not $orgMap.ContainsKey($orgId)) {
                    $orgMap[$orgId] = [PSCustomObject]@{
                        "organization_id"           = $coll.organization.id
                        "organization_name"         = $coll.organization.name
                        "collector_count"           = 1
                        "collector_ids"             = @($coll.id)
                        "collector_names"           = @($coll.display_name)
                        "number_of_managed_devices" = $collectorManagedCount
                    }
                }
                else {
                    $orgMap[$orgId].collector_count++
                    $orgMap[$orgId].collector_ids += $coll.id
                    $orgMap[$orgId].collector_names += $coll.display_name
                    $orgMap[$orgId].number_of_managed_devices += $collectorManagedCount
                }
            }
        }
        
        # Convert to flat data for Excel
        foreach ($orgId in $orgMap.Keys) {
            $org = $orgMap[$orgId]
            $organizationsData += [PSCustomObject]@{
                "organization_id"           = $org.organization_id
                "organization_name"         = $org.organization_name
                "collector_count"           = $org.collector_count
                "number_of_managed_devices" = $org.number_of_managed_devices
                "collector_ids"             = ($org.collector_ids -join "`n")
                "collector_names"           = ($org.collector_names -join "`n")
            }
        }
        
        $orgMsg = "[OK] Extracted $($organizationsData.Count) organization(s)"
        Write-Host $orgMsg -ForegroundColor Green
        $orgMsg | Out-File -FilePath $logFile -Append
    }
    
    # STEP 5: Extract Collectors data
    if ($objectTypes -contains "collectors") {
        $step5Msg = "`n[STEP 5] Extracting Collectors data..."
        Write-Host $step5Msg -ForegroundColor Cyan
        $step5Msg | Out-File -FilePath $logFile -Append
        
        foreach ($coll in $collectors) {
            $collectorRow = [PSCustomObject]@{
                "collector_id"              = $coll.id
                "display_name"              = $coll.display_name
                "number_of_managed_devices" = $managedDeviceCountByCollector[$coll.id]
                "status"                    = if ($coll.status) { $coll.status.value } else { "" }
                "status_last_change"        = if ($coll.status) { $coll.status.last_change } else { "" }
                "creation_time"             = $coll.creation_time
                "timezone"                  = $coll.timezone
                "agent_version"             = if ($coll.version) { $coll.version.agent } else { "" }
                "package_version"           = if ($coll.version) { $coll.version.package } else { "" }
                "organization_id"           = if ($coll.organization) { $coll.organization.id } else { "" }
                "organization_name"         = if ($coll.organization) { $coll.organization.name } else { "" }
                "licence_id"                = if ($coll.licence) { $coll.licence.id } else { "" }
                "licence_code"              = if ($coll.licence) { $coll.licence.code } else { "" }
                "licence_type"              = if ($coll.licence) { $coll.licence.type } else { "" }
                "licence_activation"        = if ($coll.licence) { $coll.licence.activation_time } else { "" }
                "bound_mac_address"         = if ($coll.licence) { $coll.licence.bound_mac_address } else { "" }
                "access_status"             = if ($coll.access_right) { $coll.access_right.status } else { "" }
                "api_enabled"               = if ($coll.access_right) { $coll.access_right.api_enabled } else { "" }
                "latitude"                  = if ($coll.location) { $coll.location.latitude } else { "" }
                "longitude"                 = if ($coll.location) { $coll.location.longitude } else { "" }
                "wan_ip"                    = if ($coll.wan_info) { $coll.wan_info.ip } else { "" }
                "wan_hostname"              = if ($coll.wan_info) { $coll.wan_info.hostname } else { "" }
            }
            $collectorsData += $collectorRow
        }
        
        $collDataMsg = "[OK] Extracted $($collectorsData.Count) collector(s)"
        Write-Host $collDataMsg -ForegroundColor Green
        $collDataMsg | Out-File -FilePath $logFile -Append
    }
    
    # STEP 6: Extract Devices data
    if ($objectTypes -contains "devices") {
        $step6Msg = "`n[STEP 6] Extracting Devices data..."
        Write-Host $step6Msg -ForegroundColor Cyan
        $step6Msg | Out-File -FilePath $logFile -Append
        
        $totalManagedDevices = 0
        $totalUnmanagedDevices = 0
        
        foreach ($coll in $collectors) {
            $collDevMsg = "  Processing collector: $($coll.display_name) (ID: $($coll.id))"
            Write-Host $collDevMsg -ForegroundColor Yellow
            $collDevMsg | Out-File -FilePath $logFile -Append
            
            # Get managed devices
            if ($deviceTypes -contains "managed") {
                $managedDevices = Get-ManagedDevices -collectorID $coll.id -collectorName $coll.display_name
                
                foreach ($device in $managedDevices) {
                    $deviceRow = [PSCustomObject]@{
                        "collector_id"       = $coll.id
                        "collector_name"     = $coll.display_name
                        "device_id"          = $device.id
                        "display_name"       = $device.display_name
                        "device_type"        = "managed"
                        "status"             = $device.status
                        "last_status_change" = $device.last_status_change
                        "ip_addresses"       = if ($device.ip_addresses) { $device.ip_addresses -join ", " } else { "" }
                        "hw_address"         = $device.hw_address
                        "vendor"             = $device.vendor
                        "model"              = $device.model
                        "protocol"           = $device.protocol
                        "type_id"            = if ($device.type) { $device.type.id } else { "" }
                        "type_label"         = if ($device.type) { $device.type.label } else { "" }
                        "type_detected_id"   = if ($device.type) { $device.type.detected_id } else { "" }
                        "snmp_status"        = $device.snmp_status
                        "importance"         = $device.importance
                        "first_seen_on"      = $device.first_seen_on
                        "agent_reachable"    = $device.agent_reachable
                        "is_excluded"        = $device.is_excluded
                        "is_jammed"          = $device.is_jammed
                        "grace_period"       = $device.grace_period
                        "organization_id"    = if ($device.organization) { $device.organization.id } else { "" }
                        "organization_name"  = if ($device.organization) { $device.organization.name } else { "" }
                        "zone"               = if ($device.details) { $device.details.zone } else { "" }
                        "room"               = if ($device.details) { $device.details.room } else { "" }
                        "serial"             = if ($device.details) { $device.details.serial } else { "" }
                        "firmware_version"   = if ($device.details) { $device.details.firmware_version } else { "" }
                        "notes"              = if ($device.details) { $device.details.notes } else { "" }
                        "host_name"          = if ($device.names) { $device.names.host } else { "" }
                        "bonjour_name"       = if ($device.names) { $device.names.bonjour } else { "" }
                        "upnp_name"          = if ($device.names) { $device.names.upnp } else { "" }
                        "netbios_name"       = if ($device.names) { $device.names.netbios } else { "" }
                        "tcp_ports"          = if ($device.open_ports -and $device.open_ports.tcp) { $device.open_ports.tcp -join ", " } else { "" }
                        "udp_ports"          = if ($device.open_ports -and $device.open_ports.udp) { $device.open_ports.udp -join ", " } else { "" }
                    }
                    $devicesData += $deviceRow
                }
                
                $totalManagedDevices += $managedDevices.Count
                $managedMsg = "    [OK] Found $($managedDevices.Count) managed device(s)"
                Write-Host $managedMsg -ForegroundColor Green
                $managedMsg | Out-File -FilePath $logFile -Append
            }
            
            # Get unmanaged devices
            if ($deviceTypes -contains "unmanaged") {
                $unmanagedDevices = Get-UnmanagedDevices -collectorID $coll.id -collectorName $coll.display_name
                
                foreach ($device in $unmanagedDevices) {
                    $deviceRow = [PSCustomObject]@{
                        "collector_id"       = $coll.id
                        "collector_name"     = $coll.display_name
                        "device_id"          = $device.id
                        "display_name"       = $device.display_name
                        "device_type"        = "unmanaged"
                        "status"             = ""
                        "last_status_change" = ""
                        "ip_addresses"       = if ($device.ip_addresses) { $device.ip_addresses -join ", " } else { "" }
                        "hw_address"         = $device.mac
                        "vendor"             = $device.vendor
                        "model"              = ""
                        "protocol"           = $device.protocol
                        "type_id"            = if ($device.type) { $device.type.id } else { "" }
                        "type_label"         = if ($device.type) { $device.type.label } else { "" }
                        "type_detected_id"   = if ($device.type) { $device.type.detected_id } else { "" }
                        "snmp_status"        = ""
                        "importance"         = ""
                        "first_seen_on"      = $device.first_seen_on
                        "agent_reachable"    = ""
                        "is_excluded"        = ""
                        "is_jammed"          = ""
                        "grace_period"       = ""
                        "organization_id"    = ""
                        "organization_name"  = ""
                        "zone"               = ""
                        "room"               = ""
                        "serial"             = ""
                        "firmware_version"   = ""
                        "notes"              = ""
                        "host_name"          = ""
                        "bonjour_name"       = ""
                        "upnp_name"          = ""
                        "netbios_name"       = ""
                        "tcp_ports"          = ""
                        "udp_ports"          = ""
                    }
                    $devicesData += $deviceRow
                }
                
                $totalUnmanagedDevices += $unmanagedDevices.Count
                $unmanagedMsg = "    [OK] Found $($unmanagedDevices.Count) unmanaged device(s)"
                Write-Host $unmanagedMsg -ForegroundColor Green
                $unmanagedMsg | Out-File -FilePath $logFile -Append
            }
        }
        
        $devSummaryMsg = "[OK] Total devices extracted: $($devicesData.Count) (Managed: $totalManagedDevices, Unmanaged: $totalUnmanagedDevices)"
        Write-Host $devSummaryMsg -ForegroundColor Green
        $devSummaryMsg | Out-File -FilePath $logFile -Append
    }
    
    # STEP 7: Determine file name and handle existing files
    $step7Msg = "`n[STEP 6] Creating Excel file..."
    Write-Host $step7Msg -ForegroundColor Cyan
    $step7Msg | Out-File -FilePath $logFile -Append
    
    # Determine base file name
    if ([string]::IsNullOrEmpty($fileName)) {
        $baseFileName = $DEFAULT_EXCEL_BASENAME
    }
    else {
        # Remove .xlsx extension if provided
        $baseFileName = $fileName -replace '\.xlsx$', ''
    }
    
    $targetFileName = "$baseFileName.xlsx"
    $targetFilePath = Join-Path $PSScriptRoot $targetFileName
    
    $fileExistsMessage = ""
    
    # Check if file exists and create timestamped version if needed
    if (Test-Path $targetFilePath) {
        $timestamp = Get-Date -Format "yyyy-MMM-dd_HH-mm-ss"
        $targetFileName = "${baseFileName}_${timestamp}.xlsx"
        $targetFilePath = Join-Path $PSScriptRoot $targetFileName
        
        $fileExistsMessage = @"

[NOTICE] The file '$baseFileName.xlsx' already exists.
         Creating new file with timestamp: $targetFileName
"@
        Write-Host $fileExistsMessage -ForegroundColor Yellow
        $fileExistsMessage | Out-File -FilePath $logFile -Append
    }
    
    $creatingMsg = "  Creating Excel file: $targetFileName"
    Write-Host $creatingMsg -ForegroundColor Cyan
    $creatingMsg | Out-File -FilePath $logFile -Append
    
    # STEP 8: Create Excel file with worksheets
    try {
        Import-Module ImportExcel -ErrorAction Stop
        
        $worksheetCount = 0
        
        # Create Organizations worksheet
        if ($objectTypes -contains "organizations" -and $organizationsData.Count -gt 0) {
            $organizationsData | Export-Excel -Path $targetFilePath -WorksheetName "Organizations" -FreezeTopRow -BoldTopRow
            $worksheetCount++
            
            # Apply formatting to Organizations worksheet
            $excel = Open-ExcelPackage -Path $targetFilePath
            $orgWorksheet = $excel.Workbook.Worksheets["Organizations"]
            
            if ($orgWorksheet) {
                $lastRow = $orgWorksheet.Dimension.Rows
                $lastCol = $orgWorksheet.Dimension.Columns
                
                # Apply Text format, vertical middle alignment, and left horizontal alignment to all cells
                for ($row = 1; $row -le $lastRow; $row++) {
                    for ($col = 1; $col -le $lastCol; $col++) {
                        $orgWorksheet.Cells[$row, $col].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                        $orgWorksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                        $orgWorksheet.Cells[$row, $col].Style.Numberformat.Format = "@"
                    }
                }
                
                # Find collector_ids and collector_names column indices
                $collectorIdsCol = 0
                $collectorNamesCol = 0
                
                for ($col = 1; $col -le $lastCol; $col++) {
                    $headerValue = $orgWorksheet.Cells[1, $col].Value
                    if ($headerValue -eq "collector_ids") {
                        $collectorIdsCol = $col
                    }
                    elseif ($headerValue -eq "collector_names") {
                        $collectorNamesCol = $col
                    }
                }
                
                # Apply wrap text to collector_ids column
                if ($collectorIdsCol -gt 0) {
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $orgWorksheet.Cells[$row, $collectorIdsCol].Style.WrapText = $true
                    }
                    # Set a reasonable column width
                    $orgWorksheet.Column($collectorIdsCol).Width = 15
                }
                
                # Apply wrap text to collector_names column
                if ($collectorNamesCol -gt 0) {
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $orgWorksheet.Cells[$row, $collectorNamesCol].Style.WrapText = $true
                    }
                    # Set a reasonable column width
                    $orgWorksheet.Column($collectorNamesCol).Width = 45
                }
                
                # AutoFit other columns
                for ($col = 1; $col -le $lastCol; $col++) {
                    if ($col -ne $collectorIdsCol -and $col -ne $collectorNamesCol) {
                        $orgWorksheet.Column($col).AutoFit()
                    }
                }
            }
            
            $excel.Save()
            $excel.Dispose()
            
            $orgWsMsg = "  [OK] Created 'Organizations' worksheet with $($organizationsData.Count) row(s)"
            Write-Host $orgWsMsg -ForegroundColor Green
            $orgWsMsg | Out-File -FilePath $logFile -Append
        }
        
        # Create Collectors worksheet
        if ($objectTypes -contains "collectors" -and $collectorsData.Count -gt 0) {
            if ($worksheetCount -eq 0) {
                $collectorsData | Export-Excel -Path $targetFilePath -WorksheetName "Collectors" -AutoSize -FreezeTopRow -BoldTopRow
            }
            else {
                $collectorsData | Export-Excel -Path $targetFilePath -WorksheetName "Collectors" -AutoSize -FreezeTopRow -BoldTopRow -Append
            }
            $worksheetCount++
            
            # Apply formatting to Collectors worksheet
            $excel = Open-ExcelPackage -Path $targetFilePath
            $collWorksheet = $excel.Workbook.Worksheets["Collectors"]
            
            if ($collWorksheet) {
                $lastRow = $collWorksheet.Dimension.Rows
                $lastCol = $collWorksheet.Dimension.Columns
                
                # Apply Text format, vertical middle alignment, and left horizontal alignment to all cells
                for ($row = 1; $row -le $lastRow; $row++) {
                    for ($col = 1; $col -le $lastCol; $col++) {
                        $collWorksheet.Cells[$row, $col].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                        $collWorksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                        $collWorksheet.Cells[$row, $col].Style.Numberformat.Format = "@"
                    }
                }
                
                # Build column name to index mapping
                $columnMap = @{}
                for ($col = 1; $col -le $lastCol; $col++) {
                    $headerName = $collWorksheet.Cells[1, $col].Value
                    if ($headerName) {
                        $columnMap[$headerName] = $col
                    }
                }
                
                # Apply hyperlinks to display_name column
                if ($columnMap.ContainsKey("display_name") -and $columnMap.ContainsKey("collector_id")) {
                    $displayNameColNum = $columnMap["display_name"]
                    $collectorIdColNum = $columnMap["collector_id"]
                    
                    $hyperlinkCount = 0
                    $hyperlinkMsg = "  [INFO] Adding hyperlinks to collector display_name column..."
                    Write-Host $hyperlinkMsg -ForegroundColor Cyan
                    $hyperlinkMsg | Out-File -FilePath $logFile -Append
                    
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $displayNameCell = $collWorksheet.Cells[$row, $displayNameColNum]
                        $collectorIdCell = $collWorksheet.Cells[$row, $collectorIdColNum]
                        
                        $displayName = $displayNameCell.Value
                        $collectorId = $collectorIdCell.Value
                        
                        if ($displayName -and $collectorId) {
                            # Build the Domotz portal URL for collector devices
                            $collectorUrl = "https://portal.domotz.com/webapp/agent/$collectorId/devices"
                            
                            # Create hyperlink using EPPlus
                            try {
                                $uri = New-Object System.Uri($collectorUrl)
                                $displayNameCell.Hyperlink = $uri
                                $displayNameCell.Value = $displayName
                                # Style the hyperlink (blue and underlined)
                                $displayNameCell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                $displayNameCell.Style.Font.UnderLine = $true
                                $hyperlinkCount++
                            }
                            catch {
                                $linkErrorMsg = "    [WARNING] Failed to add hyperlink for row $row`: $_"
                                Write-Host $linkErrorMsg -ForegroundColor Yellow
                                $linkErrorMsg | Out-File -FilePath $logFile -Append
                            }
                        }
                    }
                    
                    $hyperlinkDoneMsg = "  [OK] Added $hyperlinkCount hyperlinks to collector names"
                    Write-Host $hyperlinkDoneMsg -ForegroundColor Green
                    $hyperlinkDoneMsg | Out-File -FilePath $logFile -Append
                }
                
                # Ensure bound_mac_address column is NOT styled as hyperlink (black text, no underline)
                if ($columnMap.ContainsKey("bound_mac_address")) {
                    $macColNum = $columnMap["bound_mac_address"]
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $macCell = $collWorksheet.Cells[$row, $macColNum]
                        $macCell.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        $macCell.Style.Font.UnderLine = $false
                    }
                }
            }
            
            $excel.Save()
            $excel.Dispose()
            
            $collWsMsg = "  [OK] Created 'Collectors' worksheet with $($collectorsData.Count) row(s)"
            Write-Host $collWsMsg -ForegroundColor Green
            $collWsMsg | Out-File -FilePath $logFile -Append
        }
        
        # Create Devices worksheet
        if ($objectTypes -contains "devices" -and $devicesData.Count -gt 0) {
            if ($worksheetCount -eq 0) {
                $devicesData | Export-Excel -Path $targetFilePath -WorksheetName "Devices" -AutoSize -FreezeTopRow -BoldTopRow
            }
            else {
                $devicesData | Export-Excel -Path $targetFilePath -WorksheetName "Devices" -AutoSize -FreezeTopRow -BoldTopRow -Append
            }
            $worksheetCount++
            
            # Apply formatting to Devices worksheet
            $excel = Open-ExcelPackage -Path $targetFilePath
            $devWorksheet = $excel.Workbook.Worksheets["Devices"]
            
            if ($devWorksheet) {
                $lastRow = $devWorksheet.Dimension.Rows
                $lastCol = $devWorksheet.Dimension.Columns
                
                # Apply Text format, vertical middle alignment, and left horizontal alignment to all cells
                for ($row = 1; $row -le $lastRow; $row++) {
                    for ($col = 1; $col -le $lastCol; $col++) {
                        $devWorksheet.Cells[$row, $col].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                        $devWorksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                        $devWorksheet.Cells[$row, $col].Style.Numberformat.Format = "@"
                    }
                }
                
                # Build column name to index mapping
                $columnMap = @{}
                for ($col = 1; $col -le $lastCol; $col++) {
                    $headerName = $devWorksheet.Cells[1, $col].Value
                    if ($headerName) {
                        $columnMap[$headerName] = $col
                    }
                }
                
                # Get device_type column index for unmanaged device styling
                $deviceTypeColNum = if ($columnMap.ContainsKey("device_type")) { $columnMap["device_type"] } else { 0 }
                
                # Define colors for unmanaged devices
                $darkGreyColor = [System.Drawing.Color]::FromArgb(96, 96, 96)  # Dark grey for text
                $lightGreyFill = [System.Drawing.Color]::FromArgb(230, 230, 230)  # Light grey for empty cell fill
                
                # Apply styling to unmanaged devices and hyperlinks to all devices
                if ($columnMap.ContainsKey("display_name") -and $columnMap.ContainsKey("collector_id") -and $columnMap.ContainsKey("device_id")) {
                    $displayNameColNum = $columnMap["display_name"]
                    $collectorIdColNum = $columnMap["collector_id"]
                    $deviceIdColNum = $columnMap["device_id"]
                    
                    $hyperlinkCount = 0
                    $unmanagedCount = 0
                    $hyperlinkMsg = "  [INFO] Adding hyperlinks and styling devices..."
                    Write-Host $hyperlinkMsg -ForegroundColor Cyan
                    $hyperlinkMsg | Out-File -FilePath $logFile -Append
                    
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $displayNameCell = $devWorksheet.Cells[$row, $displayNameColNum]
                        $collectorIdCell = $devWorksheet.Cells[$row, $collectorIdColNum]
                        $deviceIdCell = $devWorksheet.Cells[$row, $deviceIdColNum]
                        
                        $displayName = $displayNameCell.Value
                        $collectorId = $collectorIdCell.Value
                        $deviceId = $deviceIdCell.Value
                        
                        # Check if this is an unmanaged device
                        $isUnmanaged = $false
                        if ($deviceTypeColNum -gt 0) {
                            $deviceTypeValue = $devWorksheet.Cells[$row, $deviceTypeColNum].Value
                            $isUnmanaged = ($deviceTypeValue -eq "unmanaged")
                        }
                        
                        # Apply unmanaged device styling
                        if ($isUnmanaged) {
                            $unmanagedCount++
                            for ($col = 1; $col -le $lastCol; $col++) {
                                $cell = $devWorksheet.Cells[$row, $col]
                                $cellValue = $cell.Value
                                
                                # Apply dark grey text color to all cells in the row
                                $cell.Style.Font.Color.SetColor($darkGreyColor)
                                
                                # Apply light grey fill to empty cells
                                if ([string]::IsNullOrWhiteSpace($cellValue)) {
                                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                    $cell.Style.Fill.BackgroundColor.SetColor($lightGreyFill)
                                }
                            }
                        }
                        
                        # Add hyperlink to display_name
                        if ($displayName -and $collectorId -and $deviceId) {
                            # Build the Domotz portal URL
                            $deviceUrl = "https://portal.domotz.com/webapp/agent/$collectorId/devices/$deviceId/details?tab=info"
                            
                            # Create hyperlink using EPPlus
                            try {
                                # Create a proper URI object and set it as hyperlink
                                $uri = New-Object System.Uri($deviceUrl)
                                $displayNameCell.Hyperlink = $uri
                                # Keep the display text as the device name
                                $displayNameCell.Value = $displayName
                                # Style the hyperlink (blue and underlined) - but darker blue for unmanaged
                                if ($isUnmanaged) {
                                    $displayNameCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(70, 70, 140))  # Dark blue-grey
                                }
                                else {
                                    $displayNameCell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                }
                                $displayNameCell.Style.Font.UnderLine = $true
                                $hyperlinkCount++
                            }
                            catch {
                                $linkErrorMsg = "    [WARNING] Failed to add hyperlink for row $row`: $_"
                                Write-Host $linkErrorMsg -ForegroundColor Yellow
                                $linkErrorMsg | Out-File -FilePath $logFile -Append
                            }
                        }
                    }
                    
                    $hyperlinkDoneMsg = "  [OK] Added $hyperlinkCount hyperlinks to device names"
                    Write-Host $hyperlinkDoneMsg -ForegroundColor Green
                    $hyperlinkDoneMsg | Out-File -FilePath $logFile -Append
                    
                    if ($unmanagedCount -gt 0) {
                        $unmanagedMsg = "  [OK] Styled $unmanagedCount unmanaged device rows (dark grey text, grey fill for empty cells)"
                        Write-Host $unmanagedMsg -ForegroundColor Green
                        $unmanagedMsg | Out-File -FilePath $logFile -Append
                    }
                }
                
                # Ensure hw_address column is NOT styled as hyperlink (restore proper text color, no underline)
                if ($columnMap.ContainsKey("hw_address")) {
                    $hwAddrColNum = $columnMap["hw_address"]
                    for ($row = 2; $row -le $lastRow; $row++) {
                        $hwCell = $devWorksheet.Cells[$row, $hwAddrColNum]
                        
                        # Check if this row is unmanaged (keep dark grey) or managed (use black)
                        $isUnmanagedRow = $false
                        if ($deviceTypeColNum -gt 0) {
                            $deviceTypeValue = $devWorksheet.Cells[$row, $deviceTypeColNum].Value
                            $isUnmanagedRow = ($deviceTypeValue -eq "unmanaged")
                        }
                        
                        if ($isUnmanagedRow) {
                            $hwCell.Style.Font.Color.SetColor($darkGreyColor)
                        }
                        else {
                            $hwCell.Style.Font.Color.SetColor([System.Drawing.Color]::Black)
                        }
                        $hwCell.Style.Font.UnderLine = $false
                    }
                }
            }
            
            $excel.Save()
            $excel.Dispose()
            
            $devWsMsg = "  [OK] Created 'Devices' worksheet with $($devicesData.Count) row(s)"
            Write-Host $devWsMsg -ForegroundColor Green
            $devWsMsg | Out-File -FilePath $logFile -Append
        }
        
        if ($worksheetCount -eq 0) {
            $noDataMsg = "[WARNING] No data was extracted. Excel file not created."
            Write-Host $noDataMsg -ForegroundColor Yellow
            $noDataMsg | Out-File -FilePath $logFile -Append
            return
        }
        
        # Apply final formatting to all worksheets: borders, filters, and set active worksheet
        $finalFormatMsg = "  [INFO] Applying final formatting (borders, filters, active worksheet)..."
        Write-Host $finalFormatMsg -ForegroundColor Cyan
        $finalFormatMsg | Out-File -FilePath $logFile -Append
        
        $excel = Open-ExcelPackage -Path $targetFilePath
        
        # Define border style
        $borderStyle = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $borderColor = [System.Drawing.Color]::Black
        
        foreach ($worksheet in $excel.Workbook.Worksheets) {
            if ($worksheet.Dimension) {
                $lastRow = $worksheet.Dimension.Rows
                $lastCol = $worksheet.Dimension.Columns
                
                # Apply borders to all cells
                $dataRange = $worksheet.Cells[1, 1, $lastRow, $lastCol]
                $dataRange.Style.Border.Top.Style = $borderStyle
                $dataRange.Style.Border.Bottom.Style = $borderStyle
                $dataRange.Style.Border.Left.Style = $borderStyle
                $dataRange.Style.Border.Right.Style = $borderStyle
                $dataRange.Style.Border.Top.Color.SetColor($borderColor)
                $dataRange.Style.Border.Bottom.Color.SetColor($borderColor)
                $dataRange.Style.Border.Left.Color.SetColor($borderColor)
                $dataRange.Style.Border.Right.Color.SetColor($borderColor)
                
                # Apply AutoFilter to the first row (header row)
                $headerRange = $worksheet.Cells[1, 1, 1, $lastCol]
                $worksheet.Cells[$headerRange.Address].AutoFilter = $true
                
                # Apply yellow fill color to the header row (first row)
                $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
            }
        }
        
        # Set the Devices worksheet as active (if it exists)
        $devicesWorksheet = $excel.Workbook.Worksheets["Devices"]
        if ($devicesWorksheet) {
            $excel.Workbook.View.ActiveTab = $excel.Workbook.Worksheets.Count - 1  # Devices is last
            $devicesWorksheet.View.TabSelected = $true
            
            # Deselect other worksheets
            foreach ($ws in $excel.Workbook.Worksheets) {
                if ($ws.Name -ne "Devices") {
                    $ws.View.TabSelected = $false
                }
            }
        }
        
        $excel.Save()
        $excel.Dispose()
        
        $formatDoneMsg = "  [OK] Applied borders, filters, and set Devices as active worksheet"
        Write-Host $formatDoneMsg -ForegroundColor Green
        $formatDoneMsg | Out-File -FilePath $logFile -Append
        
        # Display summary
        $summaryMsg = @"

================================================================================
                          EXTRACTION SUMMARY                                    
================================================================================

Collectors Processed: $($collectors.Count)
"@
        
        if ($objectTypes -contains "organizations") {
            $summaryMsg += "`nOrganizations: $($organizationsData.Count)"
        }
        else {
            $summaryMsg += "`nOrganizations: Skipped (collector filter applied)"
        }
        $summaryMsg += "`nCollectors: $($collectorsData.Count)"
        $summaryMsg += "`nDevices: $($devicesData.Count) (Managed: $totalManagedDevices, Unmanaged: $totalUnmanagedDevices)"
        
        $summaryMsg += @"

Output File: $targetFileName
Worksheets Created: $worksheetCount

$fileExistsMessage

================================================================================
"@
        Write-Host $summaryMsg -ForegroundColor Green
        $summaryMsg | Out-File -FilePath $logFile -Append
        
        # Open the Excel file
        Open-Excel -fileName $targetFileName
    }
    catch {
        $errorMsg = "ERROR: Failed to create Excel file - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw $_
    }
}

# Check if help parameter is provided
if ($help -or $script:showHelpOnly) {
    Show-Help
}

# Check for invalid/unknown parameters
if ($script:hasInvalidParams) {
    $errorHeader = @"

================================================================================
                                    ERROR                                       
================================================================================
"@
    Write-Host $errorHeader -ForegroundColor Red
    
    $errorMessage = "Unknown parameter(s) detected: $($script:invalidParamsList -join ', ')"
    Write-Host $errorMessage -ForegroundColor Red
    Write-Host ""
    
    $validParamsMsg = "Valid parameters are:"
    Write-Host $validParamsMsg -ForegroundColor Yellow
    Write-Host "  -operation      : extract, list_collectors" -ForegroundColor Cyan
    Write-Host "  -collector_ids  : comma-separated list of collector IDs" -ForegroundColor Cyan
    Write-Host "  -device-type    : managed, unmanaged, or managed,unmanaged" -ForegroundColor Cyan
    Write-Host "  -filename       : custom output Excel filename" -ForegroundColor Cyan
    Write-Host "  -debug          : enable detailed logging" -ForegroundColor Cyan
    Write-Host "  -help (-h, -?)  : show help" -ForegroundColor Cyan
    Write-Host ""
    
    Show-Help
}

# Set default values for parameters
if ([string]::IsNullOrEmpty($operation)) {
    $operation = "extract"
}

if ([string]::IsNullOrEmpty($device_type)) {
    $device_type = "managed,unmanaged"
}

# Validate operation value
$validOperations = @("extract", "list_collectors")
if ($operation -notin $validOperations) {
    $errorHeader = @"

================================================================================
                                    ERROR                                       
================================================================================
"@
    Write-Host $errorHeader -ForegroundColor Red
    
    $errorMessage = "Invalid operation: '$operation'"
    Write-Host $errorMessage -ForegroundColor Red
    Write-Host ""
    
    $validOpsMsg = "Valid operations are:"
    Write-Host $validOpsMsg -ForegroundColor Yellow
    foreach ($validOp in $validOperations) {
        Write-Host "  - $validOp" -ForegroundColor Cyan
    }
    Write-Host ""
    
    # Log the error
    $logMessage = @"

$errorHeader
$errorMessage

$validOpsMsg
$(foreach ($validOp in $validOperations) { "  - $validOp" })

"@
    $logMessage | Out-File -FilePath $logFile -Append
    
    Write-Host "Showing usage instructions..." -ForegroundColor Yellow
    Write-Host ""
    Show-Help
}

# Validate device-type values
if (-not [string]::IsNullOrEmpty($device_type)) {
    $deviceTypeArray = $device_type.ToLower() -split ',' | ForEach-Object { $_.Trim() }
    $validDeviceTypes = @("managed", "unmanaged")
    
    foreach ($devType in $deviceTypeArray) {
        if ($devType -notin $validDeviceTypes) {
            $errorHeader = @"

================================================================================
                                    ERROR                                       
================================================================================
"@
            Write-Host $errorHeader -ForegroundColor Red
            
            $errorMessage = "Invalid device-type: '$devType'"
            Write-Host $errorMessage -ForegroundColor Red
            Write-Host ""
            
            $validTypesMsg = "Valid device-type values are:"
            Write-Host $validTypesMsg -ForegroundColor Yellow
            foreach ($validType in $validDeviceTypes) {
                Write-Host "  - $validType" -ForegroundColor Cyan
            }
            Write-Host ""
            
            Show-Help
        }
    }
}

# Initialize log file with header for this execution
Write-LogHeader

# Main execution logic based on operation
switch ($operation) {
    "list_collectors" {
        # List all available collectors (suppress return value output)
        $null = List-Collectors -numbered $true
    }
    "extract" {
        # Extract data and create Excel
        Extract-Data -collectorIds $collector_ids -deviceType $device_type -fileName $filename
    }
    default {
        Write-Host "ERROR: Invalid operation specified!" -ForegroundColor Red
        Show-Usage
    }
}

$logMessage = "`nLOG FILE: $logFile"
Write-Host $logMessage
$logMessage | Out-File -FilePath $logFile -Append

