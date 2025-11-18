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
# This script associates the same Domotz Custom script to multiple devices across multiple collectors.
#
# EXCEL FILE FORMAT:
# The script expects an Excel file (.xlsx) or CSV file with at least the following columns:
#   - collector_id: The Domotz collector/agent ID
#   - ip_address: The IP address of the device
# Additional columns are allowed and will be displayed in the troubleshooting output but won't affect the script operation.
#
# EXCEL FILE LOCATION:
# By default, the script looks for the default Excel file in the same directory as the script.
# You can specify a different file using the -filename parameter.
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
    [string]$operation,
    [string]$script_name,
    [string]$filename,
    [string]$collector_ids,
    [switch]$debug,
    [Alias("h", "?")]
    [switch]$help
)

# Check for help arguments (support both / and - prefixes)
if ($args -contains "/help" -or $args -contains "/?" -or $args -contains "/h") {
    # Set a flag to show help after functions are defined
    $script:showHelpOnly = $true
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

# Function to list only custom drivers/scripts
function List-Scripts {
    param (
        [bool]$numbered = $false,
        [bool]$silent = $false
    )
    
    # PART 1: List Custom Drivers/Scripts
    try {
        $endpoint = "$baseURL/custom-driver"
        $headers = @{
            "X-Api-Key"    = $apiKey
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        
        if ($response.Count -eq 0) {
            if (-not $silent) {
                $noScriptsMsg = "No custom drivers/scripts found in your Domotz account."
                Write-Host $noScriptsMsg -ForegroundColor Yellow
                $noScriptsMsg | Out-File -FilePath $logFile -Append
            }
            return @()
        }
        else {
            $sortedScripts = $response | Sort-Object name
            
            if (-not $silent) {
                $headerMsg = @"

================================================================================
AVAILABLE CUSTOM DRIVERS/SCRIPTS
================================================================================
"@
                Write-Host $headerMsg -ForegroundColor Green
                $headerMsg | Out-File -FilePath $logFile -Append
                
                $index = 1
                
                foreach ($script in $sortedScripts) {
                    if ($numbered) {
                        $scriptLine = "  [$index] '$($script.name)' (ID: $($script.id))"
                    }
                    else {
                        $scriptLine = "  - '$($script.name)' (ID: $($script.id))"
                    }
                    Write-Host $scriptLine
                    $scriptLine | Out-File -FilePath $logFile -Append
                    $index++
                }
                
                $summaryMsg = "`nTotal: $($response.Count) custom driver(s)/script(s) found."
                Write-Host $summaryMsg -ForegroundColor Yellow
                $summaryMsg | Out-File -FilePath $logFile -Append
            }
            
            return $sortedScripts
        }
    }
    catch {
        $errorMsg = "ERROR: Failed to retrieve custom drivers - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        return @()
    }
}

# Function to list only collectors/agents
function List-Collectors {
    param (
        [bool]$numbered = $false,
        [bool]$silent = $false
    )
    
    # PART 2: List Collectors/Agents
    try {
        $agentEndpoint = "$baseURL/agent"
        $agentHeaders = @{
            "X-Api-Key"    = $apiKey
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
        }
        
        $agents = Invoke-RestMethod -Uri $agentEndpoint -Method Get -Headers $agentHeaders
        
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
    catch {
        $agentErrorMsg = "`nERROR: Failed to retrieve collectors/agents - $_"
        Write-Host $agentErrorMsg -ForegroundColor Red
        $agentErrorMsg | Out-File -FilePath $logFile -Append
        return @()
    }
}

# Function to list all available custom scripts and collectors (list-scripts-parameters operation)
function List-Scripts-Parameters {
    $message = "`n=== Listing Available Custom Drivers/Scripts and Collectors ===`n"
    Write-Host $message -ForegroundColor Magenta
    $message | Out-File -FilePath $logFile -Append
    
    # Get scripts data
    try {
        $endpoint = "$baseURL/custom-driver"
        $headers = @{
            "X-Api-Key"    = $apiKey
            "Content-Type" = "application/json"
        }
        
        $scripts = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        
        if ($scripts.Count -eq 0) {
            $noScriptsMsg = "No custom drivers/scripts found in your Domotz account."
            Write-Host $noScriptsMsg -ForegroundColor Yellow
            $noScriptsMsg | Out-File -FilePath $logFile -Append
        }
        else {
            $headerMsg = @"

================================================================================
AVAILABLE CUSTOM DRIVERS/SCRIPTS
================================================================================
"@
            Write-Host $headerMsg -ForegroundColor Green
            $headerMsg | Out-File -FilePath $logFile -Append
            
            $sortedScripts = $scripts | Sort-Object name
            foreach ($script in $sortedScripts) {
                $scriptLine = "  - $($script.name)"
                Write-Host $scriptLine
                $scriptLine | Out-File -FilePath $logFile -Append
            }
            
            Write-Host ""
        }
    }
    catch {
        $errorMsg = "ERROR: Failed to retrieve custom drivers - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
    }
    
    # Get collectors data
    try {
        $agentEndpoint = "$baseURL/agent"
        $agentHeaders = @{
            "X-Api-Key"    = $apiKey
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
        }
        
        $collectors = Invoke-RestMethod -Uri $agentEndpoint -Method Get -Headers $agentHeaders
        
        if ($collectors.Count -eq 0) {
            $noAgentsMsg = "`nNo collectors/agents found in your Domotz account."
            Write-Host $noAgentsMsg -ForegroundColor Yellow
            $noAgentsMsg | Out-File -FilePath $logFile -Append
        }
        else {
            $agentHeaderMsg = @"

================================================================================
AVAILABLE COLLECTORS/AGENTS
================================================================================
"@
            Write-Host $agentHeaderMsg -ForegroundColor Green
            $agentHeaderMsg | Out-File -FilePath $logFile -Append
            
            $sortedCollectors = $collectors | Sort-Object display_name
            foreach ($collector in $sortedCollectors) {
                $collectorLine = "  - $($collector.display_name)"
                Write-Host $collectorLine
                $collectorLine | Out-File -FilePath $logFile -Append
            }
            
            Write-Host ""
        }
    }
    catch {
        $agentErrorMsg = "`nERROR: Failed to retrieve collectors/agents - $_"
        Write-Host $agentErrorMsg -ForegroundColor Red
        $agentErrorMsg | Out-File -FilePath $logFile -Append
    }
    
    # Final separator
    $finalMsg = @"

================================================================================
"@
    Write-Host $finalMsg
    $finalMsg | Out-File -FilePath $logFile -Append
}

# Function to show help (usage only, no interactive workflow)
function Show-Help {
    $usageMessage = @"
================================================================================
        DOMOTZ AUTOMATION SCRIPTS - BULK APPLY TOOL
================================================================================

USAGE: .\$PS_SCRIPT_NAME.ps1 -operation <operation_type> [additional parameters]

================================================================================
OPERATION TYPES
================================================================================

    list-scripts-parameters : List all available custom drivers/scripts and collectors
                              No additional parameters required

    create-excel : Create a new Excel file with device data from specified collectors
                   Required: -script_name <script_name> -collector_ids <comma_separated_ids>
                   Optional: -filename <excel_file_name>
    
    bulk-apply   : Apply a Domotz custom script to multiple devices listed in Excel
                   Required: -script_name <script_name>
                   Optional: -filename <excel_file_name>
                   Optional: -debug (enables detailed API request logging)

================================================================================
WORKFLOW EXAMPLES
================================================================================

STEP 1: Create Excel file with devices from your collectors
--------
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759,312189"

With custom filename:
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759" -filename "poly_devices"

STEP 2: At the end of this previous operation the created file is opened. Edit the Excel file
--------

IMPORTANT: Fill in ALL required fields (marked in RED in the Excel header):
- username, password (if script requires credentials)
- Script parameters (e.g., client_id, client_secret, etc.)
- sample_period (must be >= minimal_sample_period)
NOTE: Rows with missing required fields will be SKIPPED during bulk-apply

STEP 3: Apply the script to all devices in the Excel file with _operation_ in ("DeleteAssociation", "Associate", "UpdateParameters"). _operation_ column is required.
--------
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring"

Or with specific file:
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring" -filename "poly_devices.xlsx"

Or with debug mode enabled:
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring" -debug

NOTE: The script processes only rows where _apply-result_ is empty or "Skipped":
- Rows with status "OK", "Error", or "Script already applied" are skipped (already processed)
- Rows with status "Skipped" are re-evaluated: if all required parameters are now
  filled, the association is attempted; otherwise, they remain skipped

STEP 4: Review results
--------
The Excel file will be updated with results in _apply-result_ and _messages_ columns:
- OK (green) = Success
- Error (red) = Failed (see _messages_ for details)
- Skipped (red) = Missing required parameters (see _messages_ for details)
- Script already applied = Device already has this script associated (from create-excel)

Fix any skipped/failed rows and re-run bulk-apply to process them.

================================================================================
"@
    Write-Host $usageMessage -ForegroundColor Yellow
    exit
}

# Function to show usage
function Show-Usage {
    $usageMessage = @"
================================================================================
        DOMOTZ AUTOMATION SCRIPTS - BULK APPLY TOOL
================================================================================

USAGE: .\$PS_SCRIPT_NAME.ps1 -operation <operation_type> [additional parameters]

================================================================================
OPERATION TYPES
================================================================================

    list-scripts-parameters : List all available custom drivers/scripts and collectors
                              No additional parameters required

    create-excel : Create a new Excel file with device data from specified collectors
                   Required: -script_name <script_name> -collector_ids <comma_separated_ids>
                   Optional: -filename <excel_file_name>
    
    bulk-apply   : Apply a Domotz custom script to multiple devices listed in Excel
                   Required: -script_name <script_name>
                   Optional: -filename <excel_file_name>
                   Optional: -debug (enables detailed API request logging)

================================================================================
WORKFLOW EXAMPLES
================================================================================

STEP 1: Create Excel file with devices from your collectors
--------
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759,312189"

With custom filename:
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759" -filename "poly_devices"

STEP 2: At the end of this previous operation the created file is opened. Edit the Excel file
--------

IMPORTANT: Fill in ALL required fields (marked in RED in the Excel header):
- username, password (if script requires credentials)
- Script parameters (e.g., client_id, client_secret, etc.)
- sample_period (must be >= minimal_sample_period)
NOTE: Rows with missing required fields will be SKIPPED during bulk-apply

STEP 3: Apply the script to all devices in the Excel file with _operation_ in ("DeleteAssociation", "Associate", "UpdateParameters"). _operation_ column is required.
--------
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring"

Or with specific file:
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring" -filename "poly_devices.xlsx"

Or with debug mode enabled:
.\$PS_SCRIPT_NAME.ps1 -operation bulk-apply -script_name "Poly Monitoring" -debug

NOTE: The script processes only rows where _apply-result_ is empty or "Skipped":
- Rows with status "OK", "Error", or "Script already applied" are skipped (already processed)
- Rows with status "Skipped" are re-evaluated: if all required parameters are now
  filled, the association is attempted; otherwise, they remain skipped

STEP 4: Review results
--------
The Excel file will be updated with results in _apply-result_ and _messages_ columns:
- OK (green) = Success
- Error (red) = Failed (see _messages_ for details)
- Skipped (red) = Missing required parameters (see _messages_ for details)
- Script already applied = Device already has this script associated (from create-excel)

Fix any skipped/failed rows and re-run bulk-apply to process them.

================================================================================
"@
    Write-Host $usageMessage -ForegroundColor Yellow
    $usageMessage | Out-File -FilePath $logFile -Append
    
    # Initialize wizard command history tracker
    $script:wizardCommandHistory = @()
    
    # Ask if user wants help creating the command
    Write-Host ""

    Write-Host "Do you want help creating the first command i.e. the create-excel command? (Y/N): " -ForegroundColor Cyan -NoNewline
    $response = Read-Host
    
    if ($response -notmatch '^[Yy]') {
        # User doesn't want help - show STEP 1 example and exit
        $step1Example = @"

STEP 1: Create Excel file with devices from your collectors
--------
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759,312189"

With custom filename:
.\$PS_SCRIPT_NAME.ps1 -operation create-excel -script_name "Poly Monitoring" -collector_ids "313759" -filename "poly_devices"

"@
        Write-Host $step1Example -ForegroundColor Yellow
        exit
    }
    
    # User wants help - now display available scripts with numbering
    Write-Host ""
    $scripts = List-Scripts -numbered $true
    
    if ($scripts.Count -eq 0) {
        Write-Host "`nNo scripts available. Cannot proceed with interactive wizard." -ForegroundColor Red
        exit
    }
    
    # Ask for script selection
    Write-Host ""
    Write-Host "Enter the INDEX (i.e. number w/o []) of the script you want to use: " -ForegroundColor Cyan -NoNewline
    $scriptNumber = Read-Host
    
    # Validate script number
    if (-not ($scriptNumber -match '^\d+$') -or [int]$scriptNumber -lt 1 -or [int]$scriptNumber -gt $scripts.Count) {
        Write-Host "`nERROR: Invalid script number. Please enter a number between 1 and $($scripts.Count)." -ForegroundColor Red
        exit
    }
    
    $selectedScript = $scripts[[int]$scriptNumber - 1]
    Write-Host "`nSelected script: '$($selectedScript.name)'" -ForegroundColor Green
    
    # List collectors with numbering
    Write-Host ""
    $collectors = List-Collectors -numbered $true
    
    if ($collectors.Count -eq 0) {
        Write-Host "`nNo collectors available. Cannot proceed." -ForegroundColor Red
        exit
    }
    
    # Ask for collector selection
    Write-Host ""
    Write-Host "Enter the INDEX (i.e. number w/o []) of the collectors you want to get devices from (comma-separated, e.g., 1,2,3):" -ForegroundColor Cyan
    $collectorNumbers = Read-Host "Collector numbers"
    
    # Parse and validate collector numbers
    $selectedCollectorIds = @()
    $collectorNumberArray = $collectorNumbers -split ',' | ForEach-Object { $_.Trim() }
    
    foreach ($num in $collectorNumberArray) {
        if (-not ($num -match '^\d+$') -or [int]$num -lt 1 -or [int]$num -gt $collectors.Count) {
            Write-Host "`nERROR: Invalid collector number '$num'. Please enter numbers between 1 and $($collectors.Count)." -ForegroundColor Red
            exit
        }
        $selectedCollectorIds += $collectors[[int]$num - 1].id
    }
    
    Write-Host "`nSelected collectors:" -ForegroundColor Green
    foreach ($num in $collectorNumberArray) {
        $collector = $collectors[[int]$num - 1]
        Write-Host "  - '$($collector.display_name)' (ID: $($collector.id))" -ForegroundColor Green
    }
    
    # Build the command
    $collectorIdsString = $selectedCollectorIds -join ','
    $psName = Split-Path -Leaf $PSCommandPath
    $command = ".\$psName -operation create-excel -script_name `"$($selectedScript.name)`" -collector_ids `"$collectorIdsString`""
    
    # Display the command
    Write-Host ""
    Write-Host "================================================================================`n" -ForegroundColor Yellow
    Write-Host "STEP 1: Create Excel file with devices from your collectors" -ForegroundColor Yellow
    Write-Host "--------" -ForegroundColor Yellow
    Write-Host $command -ForegroundColor Yellow
    Write-Host ""
    Write-Host "================================================================================`n" -ForegroundColor Yellow
    
    # Ask if user wants to run the command
    Write-Host "Do you want to run this command now to create the Excel file? (Y/N): " -ForegroundColor Cyan -NoNewline
    $runResponse = Read-Host
    
    if ($runResponse -match '^[Yy]') {
        Write-Host "`nExecuting command..." -ForegroundColor Green
        Write-Host ""
        
        # Track command in wizard history
        $script:wizardCommandHistory += $command
        
        # Set the script variables to execute the command
        $script:operation = "create-excel"
        $script:script_name = $selectedScript.name
        $script:collector_ids = $collectorIdsString
        
        # Continue with execution (don't exit)
        return
    }
    else {
        Write-Host "`nCommand not executed. You can copy and run it manually." -ForegroundColor Yellow
        exit
    }
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

# Check if help parameter is provided
if ($help -or $script:showHelpOnly) {
    Show-Help
}

# Validate operation parameter
if ([string]::IsNullOrEmpty($operation)) {
    Show-Usage
    # After Show-Usage, check if operation was set by the wizard
    if ([string]::IsNullOrEmpty($operation)) {
        # User chose not to run command, exit
        exit
    }
}

# Validate operation value
$validOperations = @("create-excel", "list-scripts-parameters", "bulk-apply")
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

# Initialize log file with header for this execution
Write-LogHeader

# Function to get device list for a collector
function Get-DeviceList {
    param (
        [string]$collectorID
    )
    
    try {
        $apiEndpoint = "$baseURL/agent/$collectorID/device"
        $headers = @{
            "Accept"    = "application/json"
            "X-Api-Key" = $apiKey
        }
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        return $response
    }
    catch {
        $errorMessage = "ERROR: Failed to get device list for Collector ID $collectorID - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return $null
    }
}

# Function to get device ID from IP
function Get-DeviceIDFromIP {
    param (
        [string]$deviceIP,
        [string]$collectorID,
        [array]$deviceList
    )
    
    $device = $deviceList | Where-Object { $_.ip_addresses -contains $deviceIP }
    if ($device) {
        $logMessage = "Mapped Device IP $deviceIP to Device ID $($device.id) on Collector ID $collectorID"
        Write-Host $logMessage
        $logMessage | Out-File -FilePath $logFile -Append
        return $device.id
    }
    
    $errorMessage = "ERROR: No device found with IP $deviceIP on Collector ID $collectorID"
    Write-Host "$errorMessage`n" -ForegroundColor Red
    "$errorMessage`n" | Out-File -FilePath $logFile -Append
    return $null
}

# Function to get script ID from script name
function Get-CustomDriverID {
    param (
        [string]$scriptName
    )
    
    $message = "`n=== Retrieving Custom Driver/Script ID ===`n"
    Write-Host $message -ForegroundColor Cyan
    $message | Out-File -FilePath $logFile -Append
    
    try {
        $apiEndpoint = "$baseURL/custom-driver"
        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
            "X-Api-Key"    = $apiKey
        }
        
        $logMessage = "Fetching custom drivers from: $apiEndpoint"
        Write-Host $logMessage
        $logMessage | Out-File -FilePath $logFile -Append
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        
        $logMessage = "Retrieved $($response.Count) custom drivers/scripts"
        Write-Host $logMessage -ForegroundColor Green
        $logMessage | Out-File -FilePath $logFile -Append
        
        # Find the script with matching name
        $matchingScript = $response | Where-Object { $_.name -eq $scriptName }
        
        if ($matchingScript) {
            $successMessage = "`n>>> FOUND: Script '$scriptName' has ID = $($matchingScript.id) <<<"
            Write-Host $successMessage -ForegroundColor Green
            $successMessage | Out-File -FilePath $logFile -Append
            
            # Display script details
            $detailsMessage = @"
Script Details:
  - ID: $($matchingScript.id)
  - Name: $($matchingScript.name)
  - Type: $($matchingScript.type)
  - Description: $($matchingScript.description)
  - Valid: $($matchingScript.is_valid)
  - Requires Credentials: $($matchingScript.requires_credentials)
  - Currently Applied to $($matchingScript.device_ids.Count) device(s)
"@
            Write-Host $detailsMessage -ForegroundColor Cyan
            $detailsMessage | Out-File -FilePath $logFile -Append
            
            return $matchingScript.id
        }
        else {
            $errorMessage = "ERROR: No custom driver/script found with name '$scriptName'"
            Write-Host $errorMessage -ForegroundColor Red
            $errorMessage | Out-File -FilePath $logFile -Append
            
            # List available scripts
            $availableMessage = "`nAvailable custom drivers/scripts:"
            Write-Host $availableMessage -ForegroundColor Yellow
            $availableMessage | Out-File -FilePath $logFile -Append
            
            foreach ($script in $response | Sort-Object name) {
                $scriptLine = "  - '$($script.name)' (ID: $($script.id))"
                Write-Host $scriptLine
                $scriptLine | Out-File -FilePath $logFile -Append
            }
            
            return $null
        }
    }
    catch {
        $errorMessage = "ERROR: Failed to retrieve custom drivers - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return $null
    }
}

# Function to open Excel file (open-excel operation)
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

# Function to create Excel file with device data (create-excel operation)
function Create-Excel {
    param (
        [string]$scriptName,
        [string]$collectorIds,
        [string]$fileName
    )
    
    $message = "`n=== Creating Excel File for Custom Script: $scriptName ===`n"
    Write-Host $message -ForegroundColor Magenta
    $message | Out-File -FilePath $logFile -Append
    
    # STEP 1: Validate collector_ids parameter
    $validateMsg = "`n[STEP 1] Validating collector_ids parameter..."
    Write-Host $validateMsg -ForegroundColor Cyan
    $validateMsg | Out-File -FilePath $logFile -Append
    
    if ([string]::IsNullOrEmpty($collectorIds)) {
        $errorMsg = "ERROR: -collector_ids parameter is mandatory for create-excel operation!"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw "Missing required parameter: collector_ids"
    }
    
    # Parse provided collector IDs
    $collectorArray = $collectorIds -split ',' | ForEach-Object { $_.Trim() }
    $providedMsg = "  Provided $($collectorArray.Count) collector ID(s): $($collectorArray -join ', ')"
    Write-Host $providedMsg -ForegroundColor Cyan
    $providedMsg | Out-File -FilePath $logFile -Append
    
    # Retrieve valid collectors from API
    $validationMsg = "  Retrieving valid collectors from API..."
    Write-Host $validationMsg -ForegroundColor Cyan
    $validationMsg | Out-File -FilePath $logFile -Append
    
    $validCollectors = List-Collectors -numbered $false -silent $true
    
    if ($validCollectors.Count -eq 0) {
        $errorMsg = "`nERROR: No collectors found in your Domotz account. Cannot proceed."
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw "No collectors available"
    }
    
    # Build a hashtable of valid collector IDs for quick lookup
    $validCollectorIds = @{}
    foreach ($collector in $validCollectors) {
        $validCollectorIds[$collector.id.ToString()] = $collector.display_name
    }
    
    # Validate each provided collector ID
    $invalidCollectors = @()
    $validatedCollectors = @()
    
    foreach ($collectorId in $collectorArray) {
        if ($validCollectorIds.ContainsKey($collectorId)) {
            $validatedCollectors += @{
                id   = $collectorId
                name = $validCollectorIds[$collectorId]
            }
        }
        else {
            $invalidCollectors += $collectorId
        }
    }
    
    # Report validation results
    if ($invalidCollectors.Count -gt 0) {
        $errorMsg = @"

================================================================================
                                    ERROR                                       
================================================================================

The following collector ID(s) are not valid:
  - $($invalidCollectors -join "`n  - ")

================================================================================
"@
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        
        # Ask if user wants to correct the command
        Write-Host ""
        Write-Host "Do you want to correct the command with valid collector IDs? (Y/N): " -ForegroundColor Cyan -NoNewline
        $correctResponse = Read-Host
        
        if ($correctResponse -notmatch '^[Yy]') {
            Write-Host "`nOperation cancelled. Please run the command again with valid collector IDs." -ForegroundColor Yellow
            Write-Host ""
            return
        }
        
        # Show collectors with numbers
        Write-Host ""
        $sortedCollectors = $validCollectors | Sort-Object display_name
        $collectorHeaderMsg = @"

================================================================================
AVAILABLE COLLECTORS/AGENTS
================================================================================
"@
        Write-Host $collectorHeaderMsg -ForegroundColor Green
        
        $index = 1
        foreach ($collector in $sortedCollectors) {
            $collectorLine = "  [$index] '$($collector.display_name)' (ID: $($collector.id))"
            Write-Host $collectorLine
            $index++
        }
        
        Write-Host ""
        Write-Host "Total: $($sortedCollectors.Count) collector(s)/agent(s) found." -ForegroundColor Yellow
        Write-Host ""
        
        # Ask for collector selection
        Write-Host "Enter the collector numbers you want to use (comma-separated, e.g., 1,2,3):" -ForegroundColor Cyan
        $collectorNumbers = Read-Host "Collector numbers"
        
        # Parse and validate collector numbers
        $selectedCollectorIds = @()
        $collectorNumberArray = $collectorNumbers -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($num in $collectorNumberArray) {
            if (-not ($num -match '^\d+$') -or [int]$num -lt 1 -or [int]$num -gt $sortedCollectors.Count) {
                Write-Host "`nERROR: Invalid collector number '$num'. Please enter numbers between 1 and $($sortedCollectors.Count)." -ForegroundColor Red
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                return
            }
            $selectedCollectorIds += $sortedCollectors[[int]$num - 1].id
        }
        
        # Build the corrected command
        $collectorIdsString = $selectedCollectorIds -join ','
        $psName = Split-Path -Leaf $PSCommandPath
        
        # Build command with proper escaping for display
        if ([string]::IsNullOrEmpty($fileName)) {
            $correctedCommand = ".\$psName -operation create-excel -script_name `"$scriptName`" -collector_ids `"$collectorIdsString`""
        }
        else {
            $correctedCommand = ".\$psName -operation create-excel -script_name `"$scriptName`" -collector_ids `"$collectorIdsString`" -filename `"$fileName`""
        }
        
        # Display the corrected command
        Write-Host ""
        Write-Host "================================================================================`n" -ForegroundColor Cyan
        Write-Host "Corrected command:" -ForegroundColor Yellow
        Write-Host $correctedCommand -ForegroundColor Yellow
        Write-Host ""
        Write-Host "================================================================================`n" -ForegroundColor Cyan
        
        # Ask if user wants to run the command
        Write-Host "Do you want to run this command now? (Y/N): " -ForegroundColor Cyan -NoNewline
        $runResponse = Read-Host
        
        if ($runResponse -notmatch '^[Yy]') {
            Write-Host "`nCommand not executed. You can copy and run it manually." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nExecuting corrected command..." -ForegroundColor Green
        Write-Host ""
        
        # Update the collectorIds parameter with corrected values
        $collectorIds = $collectorIdsString
        $collectorArray = $collectorIds -split ',' | ForEach-Object { $_.Trim() }
        
        # Rebuild validated collectors list with corrected IDs
        $validatedCollectors = @()
        foreach ($collectorId in $collectorArray) {
            $validatedCollectors += @{
                id   = $collectorId
                name = $validCollectorIds[$collectorId]
            }
        }
        
        # Display validated collectors
        $successMsg = "`n[OK] All $($validatedCollectors.Count) collector ID(s) validated successfully:"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
        
        foreach ($validated in $validatedCollectors) {
            $validMsg = "  - ID: $($validated.id) - '$($validated.name)'"
            Write-Host $validMsg -ForegroundColor Green
            $validMsg | Out-File -FilePath $logFile -Append
        }
    }
    else {
        # Display validated collectors (original path when all IDs are valid)
        $successMsg = "`n[OK] All $($validatedCollectors.Count) collector ID(s) validated successfully:"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
        
        foreach ($validated in $validatedCollectors) {
            $validMsg = "  - ID: $($validated.id) - '$($validated.name)'"
            Write-Host $validMsg -ForegroundColor Green
            $validMsg | Out-File -FilePath $logFile -Append
        }
    }
    
    # STEP 2: Get the script ID from the script name
    $scriptIdMsg = "`n[STEP 2] Validating custom script name..."
    Write-Host $scriptIdMsg -ForegroundColor Cyan
    $scriptIdMsg | Out-File -FilePath $logFile -Append
    
    # Retrieve all scripts to validate
    $providedScriptMsg = "  Provided script name: '$scriptName'"
    Write-Host $providedScriptMsg -ForegroundColor Cyan
    $providedScriptMsg | Out-File -FilePath $logFile -Append
    
    $validationScriptMsg = "  Retrieving valid scripts from API..."
    Write-Host $validationScriptMsg -ForegroundColor Cyan
    $validationScriptMsg | Out-File -FilePath $logFile -Append
    
    $validScripts = List-Scripts -numbered $false -silent $true
    
    if ($validScripts.Count -eq 0) {
        $errorMsg = "`nERROR: No scripts found in your Domotz account. Cannot proceed."
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        return
    }
    
    # Find matching script
    $matchingScript = $validScripts | Where-Object { $_.name -eq $scriptName }
    
    if (-not $matchingScript) {
        # Script name is invalid - offer interactive correction
        $errorMsg = @"

================================================================================
                                    ERROR                                       
================================================================================

The script name '$scriptName' is not valid.

================================================================================
"@
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        
        # Ask if user wants to correct the command
        Write-Host ""
        Write-Host "Do you want to correct the command with a valid script name? (Y/N): " -ForegroundColor Cyan -NoNewline
        $correctResponse = Read-Host
        
        if ($correctResponse -notmatch '^[Yy]') {
            Write-Host "`nOperation cancelled. Please run the command again with a valid script name." -ForegroundColor Yellow
            Write-Host ""
            return
        }
        
        # Show scripts with numbers
        Write-Host ""
        $sortedScripts = $validScripts | Sort-Object name
        $scriptHeaderMsg = @"

================================================================================
AVAILABLE CUSTOM DRIVERS/SCRIPTS
================================================================================
"@
        Write-Host $scriptHeaderMsg -ForegroundColor Green
        
        $index = 1
        foreach ($script in $sortedScripts) {
            $scriptLine = "  [$index] '$($script.name)' (ID: $($script.id))"
            Write-Host $scriptLine
            $index++
        }
        
        Write-Host ""
        Write-Host "Total: $($sortedScripts.Count) custom driver(s)/script(s) found." -ForegroundColor Yellow
        Write-Host ""
        
        # Ask for script selection
        Write-Host "Enter the INDEX (w/o []) of the script you want to use:" -ForegroundColor Cyan
        $scriptNumber = Read-Host "Script number"
        
        # Validate script number
        if (-not ($scriptNumber -match '^\d+$') -or [int]$scriptNumber -lt 1 -or [int]$scriptNumber -gt $sortedScripts.Count) {
            Write-Host "`nERROR: Invalid script number. Please enter a number between 1 and $($sortedScripts.Count)." -ForegroundColor Red
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }
        
        $selectedScript = $sortedScripts[[int]$scriptNumber - 1]
        $psName = Split-Path -Leaf $PSCommandPath
        
        # Build the corrected command
        # Build command with proper escaping for display
        if ([string]::IsNullOrEmpty($fileName)) {
            $correctedCommand = ".\$psName -operation create-excel -script_name `"$($selectedScript.name)`" -collector_ids `"$collectorIds`""
        }
        else {
            $correctedCommand = ".\$psName -operation create-excel -script_name `"$($selectedScript.name)`" -collector_ids `"$collectorIds`" -filename `"$fileName`""
        }
        
        # Display the corrected command
        Write-Host ""
        Write-Host "================================================================================`n" -ForegroundColor Cyan
        Write-Host "Corrected command:" -ForegroundColor Yellow
        Write-Host $correctedCommand -ForegroundColor Yellow
        Write-Host ""
        Write-Host "================================================================================`n" -ForegroundColor Cyan
        
        # Ask if user wants to run the command
        Write-Host "Do you want to run this command now? (Y/N): " -ForegroundColor Cyan -NoNewline
        $runResponse = Read-Host
        
        if ($runResponse -notmatch '^[Yy]') {
            Write-Host "`nCommand not executed. You can copy and run it manually." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nExecuting corrected command..." -ForegroundColor Green
        Write-Host ""
        
        # Update the scriptName parameter with corrected value
        $scriptName = $selectedScript.name
        $matchingScript = $selectedScript
    }
    
    # Display validated script
    $scriptID = $matchingScript.id
    $successMsg = "`n[OK] Script '$scriptName' validated successfully (ID: $scriptID)"
    Write-Host $successMsg -ForegroundColor Green
    $successMsg | Out-File -FilePath $logFile -Append
    
    # STEP 3: Get custom driver details (parameters and requires_credentials)
    $detailsMsg = "`n[STEP 3] Retrieving custom driver details..."
    Write-Host $detailsMsg -ForegroundColor Cyan
    $detailsMsg | Out-File -FilePath $logFile -Append
    
    try {
        $endpoint = "$baseURL/custom-driver/$scriptID"
        $headers = @{
            "X-Api-Key"    = $apiKey
            "Content-Type" = "application/json"
        }
        
        $customDriverDetails = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers
        
        $successMsg = "[OK] Retrieved custom driver details successfully"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
        
        $detailInfo = @"
  - ID: $($customDriverDetails.id)
  - Name: $($customDriverDetails.name)
  - Requires Credentials: $($customDriverDetails.requires_credentials)
  - Parameter Count: $($customDriverDetails.parameters.Count)
"@
        Write-Host $detailInfo
        $detailInfo | Out-File -FilePath $logFile -Append
        
        # Extract parameter names with value types
        $parameterNames = @()
        $parameterNamesWithType = @()
        if ($customDriverDetails.parameters -and $customDriverDetails.parameters.Count -gt 0) {
            foreach ($param in $customDriverDetails.parameters) {
                $parameterNames += $param.name
                $parameterNamesWithType += "$($param.name) ($($param.value_type))"
            }
            $paramMsg = "  - Parameters: $($parameterNamesWithType -join ', ')"
            Write-Host $paramMsg
            $paramMsg | Out-File -FilePath $logFile -Append
        }
        else {
            $noParamMsg = "  - No parameters defined for this script"
            Write-Host $noParamMsg -ForegroundColor Yellow
            $noParamMsg | Out-File -FilePath $logFile -Append
        }
    }
    catch {
        $errorMsg = "ERROR: Failed to retrieve custom driver details - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw $_
    }
    
    # STEP 4: Build Excel header columns in the specified order
    $headerMsg = "`n[STEP 4] Building Excel header columns..."
    Write-Host $headerMsg -ForegroundColor Cyan
    $headerMsg | Out-File -FilePath $logFile -Append
    
    $excelHeaders = @()
    
    # Standard columns (some in red)
    $excelHeaders += "_collector_name_"
    $excelHeaders += "collector_id"  # red text
    $excelHeaders += "_device_display_name_"
    $excelHeaders += "ip_address"  # red text
    $excelHeaders += "_device_id_"
    
    # Add username and password if required
    if ($customDriverDetails.requires_credentials -eq $true) {
        $excelHeaders += "username"  # red text
        $excelHeaders += "password"  # red text
        $credHeaderMsg = "  [INFO] Adding username and password columns (requires_credentials: true)"
        Write-Host $credHeaderMsg -ForegroundColor Yellow
        $credHeaderMsg | Out-File -FilePath $logFile -Append
    }
    
    # Add script parameters (red text) with value_type
    foreach ($paramNameWithType in $parameterNamesWithType) {
        $excelHeaders += $paramNameWithType  # red text
    }
    
    # Add sample_period and minimal_sample_period
    $excelHeaders += "sample_period"  # red text
    $excelHeaders += "_minimal_sample_period_"
    
    # Add result tracking columns
    $excelHeaders += "_apply-result_"
    $excelHeaders += "_messages_"
    
    # Add additional device information columns
    $excelHeaders += "_device-status_"
    $excelHeaders += "_vendor_"
    $excelHeaders += "_model_"
    $excelHeaders += "_room_"
    $excelHeaders += "_zone_"
    $excelHeaders += "_serial_"
    $excelHeaders += "_snmp_status_"
    $excelHeaders += "_type_"
    $excelHeaders += "_hw_address_"
    
    $headerListMsg = "  [OK] Created header with $($excelHeaders.Count) columns"
    Write-Host $headerListMsg -ForegroundColor Green
    $headerListMsg | Out-File -FilePath $logFile -Append
    
    # STEP 5: Collect device data from all collectors
    $devicesMsg = "`n[STEP 5] Collecting device data from collectors..."
    Write-Host $devicesMsg -ForegroundColor Cyan
    $devicesMsg | Out-File -FilePath $logFile -Append
    
    $allDeviceData = @()
    $totalDeviceCount = 0
    
    foreach ($collectorId in $collectorArray) {
        $collectorMsg = "`n  Processing Collector ID: $collectorId"
        Write-Host $collectorMsg -ForegroundColor Yellow
        $collectorMsg | Out-File -FilePath $logFile -Append
        
        try {
            # Get collector details for name
            $collectorEndpoint = "$baseURL/agent/$collectorId"
            $collectorHeaders = @{
                "X-Api-Key"    = $apiKey
                "Content-Type" = "application/json"
            }
            $collectorDetails = Invoke-RestMethod -Uri $collectorEndpoint -Method Get -Headers $collectorHeaders
            $collectorName = $collectorDetails.display_name
            
            # Get devices for this collector
            $deviceList = Get-DeviceList -collectorID $collectorId
            
            if ($deviceList.Count -eq 0) {
                $noDevMsg = "    [WARNING] No devices found for Collector ID: $collectorId"
                Write-Host $noDevMsg -ForegroundColor Yellow
                $noDevMsg | Out-File -FilePath $logFile -Append
                continue
            }
            
            $foundMsg = "    [OK] Found $($deviceList.Count) device(s) in collector '$collectorName'"
            Write-Host $foundMsg -ForegroundColor Green
            $foundMsg | Out-File -FilePath $logFile -Append
            
            # Fetch existing script associations for this collector
            $assocMsg = "    [INFO] Fetching existing script associations for collector..."
            Write-Host $assocMsg -ForegroundColor Cyan
            $assocMsg | Out-File -FilePath $logFile -Append
            
            $deviceAssociations = @{}
            try {
                $associationEndpoint = "$baseURL/custom-driver/agent/$collectorId/association"
                $associationHeaders = @{
                    "X-Api-Key"    = $apiKey
                    "Content-Type" = "application/json"
                }
                $allAssociations = Invoke-RestMethod -Uri $associationEndpoint -Method Get -Headers $associationHeaders
                
                # Filter associations for this specific script and create lookup map
                $matchingAssociations = $allAssociations | Where-Object { $_.custom_driver_id -eq $customDriverDetails.id }
                
                foreach ($assoc in $matchingAssociations) {
                    $deviceAssociations[$assoc.device_id] = $assoc
                }
                
                if ($matchingAssociations.Count -gt 0) {
                    $assocFoundMsg = "    [OK] Found $($matchingAssociations.Count) existing association(s) for script '$($customDriverDetails.name)'"
                    Write-Host $assocFoundMsg -ForegroundColor Green
                    $assocFoundMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    $noAssocMsg = "    [INFO] No existing associations found for this script"
                    Write-Host $noAssocMsg -ForegroundColor Gray
                    $noAssocMsg | Out-File -FilePath $logFile -Append
                }
            }
            catch {
                $assocErrorMsg = "    [WARNING] Failed to fetch associations (will continue with empty values): $_"
                Write-Host $assocErrorMsg -ForegroundColor Yellow
                $assocErrorMsg | Out-File -FilePath $logFile -Append
            }
            
            # Build row data for each device
            foreach ($device in $deviceList) {
                $deviceRow = [ordered]@{}
                
                # Populate standard columns
                $deviceRow["_collector_name_"] = $collectorName
                $deviceRow["collector_id"] = $collectorId
                $deviceRow["_device_display_name_"] = $device.display_name
                # Store IP address as-is, will format as text in Excel
                $deviceRow["ip_address"] = if ($device.ip_addresses -and $device.ip_addresses.Count -gt 0) { $device.ip_addresses[0] } else { "" }
                $deviceRow["_device_id_"] = $device.id
                $deviceRow["_operation_"] = ""
                
                # Check if this device has an existing association
                $existingAssociation = $null
                if ($deviceAssociations.ContainsKey($device.id)) {
                    $existingAssociation = $deviceAssociations[$device.id]
                }
                
                # Add username/password if required
                if ($customDriverDetails.requires_credentials -eq $true) {
                    $deviceRow["username"] = ""
                    $deviceRow["password"] = ""
                    # Note: credentials are not returned by the association API, so we leave them empty
                }
                
                # Add fields for script parameters (with value_type in column name)
                # Populate with existing values if association exists
                foreach ($paramNameWithType in $parameterNamesWithType) {
                    # Parse parameter name from "name (TYPE)" format
                    if ($paramNameWithType -match '^(.+?)\s*\(([^)]+)\)\s*$') {
                        $paramName = $matches[1].Trim()
                        $paramType = $matches[2].Trim()
                        
                        # Check if existing association has this parameter
                        $paramValue = ""
                        if ($existingAssociation -and $existingAssociation.parameters) {
                            $matchingParam = $existingAssociation.parameters | Where-Object { $_.name -eq $paramName }
                            if ($matchingParam) {
                                # Format value based on type
                                if ($paramType -eq "LIST") {
                                    # Convert array to JSON format for Excel
                                    if ($matchingParam.value -is [System.Array] -and $matchingParam.value.Count -eq 0) {
                                        # Empty array - explicitly format as []
                                        $paramValue = "[]"
                                    }
                                    else {
                                        $paramValue = ($matchingParam.value | ConvertTo-Json -Compress)
                                    }
                                }
                                elseif ($paramType -eq "SECRET_TEXT") {
                                    # Secret values are masked in API response, leave empty for user to fill
                                    $paramValue = ""
                                }
                                else {
                                    # String or other types
                                    $paramValue = $matchingParam.value
                                }
                            }
                        }
                        $deviceRow[$paramNameWithType] = $paramValue
                    }
                    else {
                        $deviceRow[$paramNameWithType] = ""
                    }
                }
                
                # Add sample_period from existing association, or empty for user to fill
                if ($existingAssociation) {
                    $deviceRow["sample_period"] = $existingAssociation.sample_period
                }
                else {
                    $deviceRow["sample_period"] = ""
                }
                $deviceRow["_minimal_sample_period_"] = $customDriverDetails.minimal_sample_period
                
                # Add result tracking columns
                # Mark devices with existing associations
                if ($existingAssociation) {
                    $deviceRow["_apply-result_"] = "Script already applied"
                    $deviceRow["_messages_"] = "Association already exists with ID: $($existingAssociation.id)"
                }
                else {
                    $deviceRow["_apply-result_"] = ""
                    $deviceRow["_messages_"] = ""
                }
                
                # Add device details
                $deviceRow["_device-status_"] = if ($device.status) { $device.status } else { "" }
                $deviceRow["_vendor_"] = if ($device.vendor) { $device.vendor } else { "" }
                $deviceRow["_model_"] = if ($device.model) { $device.model } else { "" }
                $deviceRow["_room_"] = if ($device.details.room) { $device.details.room } else { "" }
                $deviceRow["_zone_"] = if ($device.details.zone) { $device.details.zone } else { "" }
                $deviceRow["_serial_"] = if ($device.details.serial) { $device.details.serial } else { "" }
                $deviceRow["_snmp_status_"] = if ($device.snmp_status) { $device.snmp_status } else { "" }
                $deviceRow["_type_"] = if ($device.type.label) { $device.type.label } else { "" }
                $deviceRow["_hw_address_"] = if ($device.hw_address) { $device.hw_address } else { "" }
                
                $allDeviceData += [PSCustomObject]$deviceRow
                $totalDeviceCount++
            }
        }
        catch {
            $errorMsg = "    [ERROR] Failed to retrieve devices from Collector ID $collectorId - $_"
            Write-Host $errorMsg -ForegroundColor Red
            $errorMsg | Out-File -FilePath $logFile -Append
        }
    }
    
    if ($allDeviceData.Count -eq 0) {
        $noDataMsg = "`n[ERROR] No device data collected. Cannot create Excel file."
        Write-Host $noDataMsg -ForegroundColor Red
        $noDataMsg | Out-File -FilePath $logFile -Append
        throw "No device data available"
    }
    
    $dataCollectedMsg = "`n[OK] Collected $totalDeviceCount device(s) total from $($collectorArray.Count) collector(s)"
    Write-Host $dataCollectedMsg -ForegroundColor Green
    $dataCollectedMsg | Out-File -FilePath $logFile -Append
    
    # STEP 6: Determine file name and handle existing files
    $fileNameMsg = "`n[STEP 6] Determining output file name..."
    Write-Host $fileNameMsg -ForegroundColor Cyan
    $fileNameMsg | Out-File -FilePath $logFile -Append
    
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
    
    # STEP 7: Create Excel file with formatting
    $excelMsg = "`n[STEP 7] Writing data to Excel file..."
    Write-Host $excelMsg -ForegroundColor Cyan
    $excelMsg | Out-File -FilePath $logFile -Append
    
    try {
        # Define columns that should have red text
        $redTextColumns = @(
            "collector_id",
            "ip_address",
            "sample_period"
        )
        
        # Add username/password to red columns if present
        if ($customDriverDetails.requires_credentials -eq $true) {
            $redTextColumns += "username"
            $redTextColumns += "password"
        }
        
        # Add all parameter names (with value_type) to red columns
        $redTextColumns += $parameterNamesWithType
        
        # Create Excel file with proper formatting
        # CRITICAL: Use NoNumberConversion to prevent ip_address from being treated as number
        $allDeviceData | Export-Excel -Path $targetFilePath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "Devices" -NoNumberConversion "ip_address"
        
        # Open the Excel package to apply custom formatting
        $excel = Open-ExcelPackage -Path $targetFilePath
        $worksheet = $excel.Workbook.Worksheets["Devices"]
        
        # Get dimensions
        $lastCol = $worksheet.Dimension.Columns
        $lastRow = $worksheet.Dimension.Rows
        
        # Create column name to index mapping
        $columnMap = @{}
        for ($col = 1; $col -le $lastCol; $col++) {
            $headerName = $worksheet.Cells[1, $col].Value
            if ($headerName) {
                $columnMap[$headerName] = $col
            }
        }
        
        # Apply RED TEXT to required input columns (entire column including header)
        $formattedCount = 0
        foreach ($requiredCol in $redTextColumns) {
            if ($columnMap.ContainsKey($requiredCol)) {
                $colNum = $columnMap[$requiredCol]
                
                # Apply red text to entire column
                for ($row = 1; $row -le $lastRow; $row++) {
                    $cell = $worksheet.Cells[$row, $colNum]
                    $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    
                    # Make header bold
                    if ($row -eq 1) {
                        $cell.Style.Font.Bold = $true
                    }
                }
                $formattedCount++
            }
        }
        
        $formatMsg = "  [INFO] Applied red text formatting to $formattedCount required columns (all cells)"
        Write-Host $formatMsg -ForegroundColor Cyan
        $formatMsg | Out-File -FilePath $logFile -Append
        
        # Apply red text formatting to _operation_ column
        if ($columnMap.ContainsKey("_operation_")) {
            $operationColNum = $columnMap["_operation_"]
            
            $operationFormatMsg = "  [INFO] Applying red text formatting to _operation_ column..."
            Write-Host $operationFormatMsg -ForegroundColor Cyan
            $operationFormatMsg | Out-File -FilePath $logFile -Append
            
            # Apply red text to entire column
            for ($row = 1; $row -le $lastRow; $row++) {
                $cell = $worksheet.Cells[$row, $operationColNum]
                $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                
                # Make header bold (row 1 only)
                if ($row -eq 1) {
                    $cell.Style.Font.Bold = $true
                }
            }
            
            $operationDoneMsg = "  [OK] Applied red text formatting to _operation_ column"
            Write-Host $operationDoneMsg -ForegroundColor Green
            $operationDoneMsg | Out-File -FilePath $logFile -Append
        }
        
        # Apply hyperlinks to _device_display_name_ column
        if ($columnMap.ContainsKey("_device_display_name_")) {
            $displayNameColNum = $columnMap["_device_display_name_"]
            $collectorIdColNum = $columnMap["collector_id"]
            $deviceIdColNum = $columnMap["_device_id_"]
            
            $hyperlinkCount = 0
            $hyperlinkMsg = "  [INFO] Adding hyperlinks to _device_display_name_ column..."
            Write-Host $hyperlinkMsg -ForegroundColor Cyan
            $hyperlinkMsg | Out-File -FilePath $logFile -Append
            
            for ($row = 2; $row -le $lastRow; $row++) {
                $displayNameCell = $worksheet.Cells[$row, $displayNameColNum]
                $collectorIdCell = $worksheet.Cells[$row, $collectorIdColNum]
                $deviceIdCell = $worksheet.Cells[$row, $deviceIdColNum]
                
                $displayName = $displayNameCell.Value
                $collectorId = $collectorIdCell.Value
                $deviceId = $deviceIdCell.Value
                
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
            
            $hyperlinkDoneMsg = "  [OK] Added $hyperlinkCount hyperlinks to device names"
            Write-Host $hyperlinkDoneMsg -ForegroundColor Green
            $hyperlinkDoneMsg | Out-File -FilePath $logFile -Append
        }
        
        # Apply formatting to "Script already applied" status cells
        if ($columnMap.ContainsKey("_apply-result_")) {
            $statusColNum = $columnMap["_apply-result_"]
            $messageColNum = if ($columnMap.ContainsKey("_messages_")) { $columnMap["_messages_"] } else { 0 }
            
            $formattedStatusCount = 0
            for ($row = 2; $row -le $lastRow; $row++) {
                $statusCell = $worksheet.Cells[$row, $statusColNum]
                $statusValue = $statusCell.Value
                
                if ($statusValue -eq "Script already applied") {
                    # Green and Bold for Script already applied
                    $statusCell.Style.Font.Bold = $true
                    $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                    $formattedStatusCount++
                    
                    # Also format the corresponding message cell if it exists
                    if ($messageColNum -gt 0) {
                        $messageCell = $worksheet.Cells[$row, $messageColNum]
                        $messageCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                    }
                }
            }
            
            if ($formattedStatusCount -gt 0) {
                $statusFormatMsg = "  [INFO] Formatted $formattedStatusCount 'Script already applied' status cells in green bold"
                Write-Host $statusFormatMsg -ForegroundColor Cyan
                $statusFormatMsg | Out-File -FilePath $logFile -Append
            }
        }
        
        # Apply data validation to _operation_ column
        if ($columnMap.ContainsKey("_operation_")) {
            $operationColNum = $columnMap["_operation_"]
            
            $validationMsg = "  [INFO] Adding data validation to _operation_ column..."
            Write-Host $validationMsg -ForegroundColor Cyan
            $validationMsg | Out-File -FilePath $logFile -Append
            
            # Apply data validation to all data rows (skip header row)
            for ($row = 2; $row -le $lastRow; $row++) {
                $cell = $worksheet.Cells[$row, $operationColNum]
                
                # Create data validation for the cell
                $validation = $cell.DataValidation.AddListDataValidation()
                $validation.ShowErrorMessage = $true
                $validation.ErrorTitle = "Invalid Operation"
                $validation.Error = "Please select a valid operation: Associate, DeleteAssociation, or UpdateParameters"
                $validation.AllowBlank = $true
                
                # Add the three allowed values
                $validation.Formula.Values.Add("Associate") | Out-Null
                $validation.Formula.Values.Add("DeleteAssociation") | Out-Null
                $validation.Formula.Values.Add("UpdateParameters") | Out-Null
            }
            
            $validationDoneMsg = "  [OK] Added data validation with dropdown list to $($lastRow - 1) cells in _operation_ column"
            Write-Host $validationDoneMsg -ForegroundColor Green
            $validationDoneMsg | Out-File -FilePath $logFile -Append
        }
        
        # Save and close
        $excel.Save()
        $excel.Dispose()
        
        $successMsg = "[OK] Excel file created successfully: $targetFileName"
        Write-Host $successMsg -ForegroundColor Green
        $successMsg | Out-File -FilePath $logFile -Append
        
        # Display final summary
        $summaryMsg = @"

================================================================================
                          CREATE-EXCEL SUMMARY                                  
================================================================================

Script Name: $scriptName
Script ID: $scriptID
Requires Credentials: $($customDriverDetails.requires_credentials)
Parameter Count: $($parameterNames.Count)
Collectors Processed: $($collectorArray.Count)
Total Devices: $totalDeviceCount
Output File: $targetFileName

$fileExistsMessage

The Excel file is ready for editing. Fill in the required fields (marked in red)
before running the bulk-apply operation.

================================================================================
"@
        Write-Host $summaryMsg -ForegroundColor Green
        $summaryMsg | Out-File -FilePath $logFile -Append
        
        # Auto-open the Excel file
        Write-Host "`nOpening Excel file for editing. Fill the " -ForegroundColor Cyan -NoNewline
        Write-Host "RED" -ForegroundColor Red -NoNewline
        Write-Host " required fields in the Excel file." -ForegroundColor Cyan
        "`nOpening Excel file for editing. Fill the RED required fields in the Excel file." | Out-File -FilePath $logFile -Append
        
        Start-Process $targetFilePath
        
        # Wait 10 seconds to allow Excel to fully open
        Start-Sleep -Seconds 10
        
        # Display next step command
        # Use the actual PowerShell script filename (not the Domotz script name)
        $psName = Split-Path -Leaf $PSCommandPath
        
        # Build message with colored "RED" word
        $nextStepHeader = @"

================================================================================
                              NEXT STEP
================================================================================

"@
        Write-Host $nextStepHeader -ForegroundColor Yellow
        $nextStepHeader | Out-File -FilePath $logFile -Append
        
        # Write the simplified message
        $nextStepMessage = @"
After you finish, run the following command to apply the script to all devices:

.\$psName -operation bulk-apply -script_name "$scriptName" -filename "$targetFileName"

================================================================================
"@
        Write-Host $nextStepMessage -ForegroundColor Yellow
        $nextStepMessage | Out-File -FilePath $logFile -Append
        
        # Ask user if they want to run the bulk-apply command now
        Write-Host ""
        Write-Host "Do you want to run the bulk-apply command now? (Y/N): " -ForegroundColor Cyan -NoNewline
        $runResponse = Read-Host
        $responseMsg = "User response: $runResponse"
        $responseMsg | Out-File -FilePath $logFile -Append
        
        if ($runResponse -match '^[Yy]') {
            $executeMsg = "`n[INFO] Executing bulk-apply operation..."
            Write-Host $executeMsg -ForegroundColor Green
            $executeMsg | Out-File -FilePath $logFile -Append
            
            # Track command in wizard history if it exists
            $bulkApplyCommand = ".\$psName -operation bulk-apply -script_name `"$scriptName`" -filename `"$targetFileName`""
            if ($null -ne $script:wizardCommandHistory) {
                $script:wizardCommandHistory += $bulkApplyCommand
            }
            
            # Call bulk-Apply-Script function
            bulk-Apply-Script -scriptName $scriptName -fileName $targetFileName
            
            # Display wizard command summary at the end (if wizard was used)
            if ($null -ne $script:wizardCommandHistory -and $script:wizardCommandHistory.Count -gt 0) {
                $wizardSummaryMsg = @"

================================================================================
                    WIZARD COMMAND HISTORY SUMMARY
================================================================================

The following commands were executed during this wizard session:

"@
                Write-Host $wizardSummaryMsg -ForegroundColor Cyan
                $wizardSummaryMsg | Out-File -FilePath $logFile -Append
                
                $commandIndex = 1
                foreach ($cmd in $script:wizardCommandHistory) {
                    # Determine command description based on operation type
                    if ($cmd -match '-operation\s+create-excel') {
                        $cmdDescription = "Create excel command executed:"
                    }
                    elseif ($cmd -match '-operation\s+bulk-apply') {
                        $cmdDescription = "Bulk Apply excel command executed:"
                    }
                    else {
                        $cmdDescription = "Command executed:"
                    }
                    
                    Write-Host $cmdDescription -ForegroundColor Cyan
                    $cmdDescription | Out-File -FilePath $logFile -Append
                    Write-Host $cmd -ForegroundColor Yellow
                    $cmd | Out-File -FilePath $logFile -Append
                    
                    # Add space between commands (except after the last one)
                    if ($commandIndex -lt $script:wizardCommandHistory.Count) {
                        Write-Host ""
                        "" | Out-File -FilePath $logFile -Append
                    }
                    
                    $commandIndex++
                }
                
                $endSummary = @"

================================================================================
"@
                Write-Host $endSummary -ForegroundColor Cyan
                $endSummary | Out-File -FilePath $logFile -Append
            }
        }
        else {
            $cancelMsg = "`n[INFO] bulk-apply operation cancelled by user. You can run it manually later."
            Write-Host $cancelMsg -ForegroundColor Yellow
            $cancelMsg | Out-File -FilePath $logFile -Append
        }
    }
    catch {
        $errorMsg = "ERROR: Failed to create Excel file - $_"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        throw $_
    }
}

# Function to apply custom script to devices (bulk-apply operation)
function bulk-Apply-Script {
    param (
        [string]$scriptName,
        [string]$fileName
    )
    
    $message = "`n=== bulk Applying Custom Script: $scriptName ===`n"
    Write-Host $message -ForegroundColor Magenta
    $message | Out-File -FilePath $logFile -Append
    
    # STEP 1: Get the script ID from the script name BEFORE processing any devices
    $scriptID = Get-CustomDriverID -scriptName $scriptName
    
    if (-not $scriptID) {
        $errorMessage = "`nERROR: Cannot proceed without valid script ID. Exiting."
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return
    }
    
    # STEP 2: Check Excel file and prepare
    # Determine which file to read
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = $DEFAULT_EXCEL_FILENAME
        
        # Check if there are other xlsx files in the folder
        $allXlsxFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.xlsx" -File | Where-Object { $_.Name -ne $DEFAULT_EXCEL_FILENAME }
        
        if ($allXlsxFiles.Count -gt 0) {
            $warningMsg = @"

================================================================================
                                  WARNING                                       
================================================================================

No -filename parameter was specified. The script will use the default file:
  $DEFAULT_EXCEL_FILENAME

However, the following other Excel files were found in the folder:
"@
            Write-Host $warningMsg -ForegroundColor Yellow
            $warningMsg | Out-File -FilePath $logFile -Append
            
            foreach ($xlsxFile in $allXlsxFiles) {
                $fileMsg = "  - $($xlsxFile.Name)"
                Write-Host $fileMsg -ForegroundColor Cyan
                $fileMsg | Out-File -FilePath $logFile -Append
            }
            
            $confirmMsg = @"

Do you want to continue using the default file '$DEFAULT_EXCEL_FILENAME'? (Y/N): 
"@
            Write-Host $confirmMsg -ForegroundColor Cyan -NoNewline
            $confirmMsg | Out-File -FilePath $logFile -Append
            
            $response = Read-Host
            $responseMsg = "User response: $response"
            $responseMsg | Out-File -FilePath $logFile -Append
            
            if ($response -notmatch '^[Yy]') {
                # User doesn't want to use default file - ask which file to use instead
                $alternativeMsg = @"

Which Excel file would you like to use instead?
Please enter the filename (e.g., 'my_devices.xlsx' or just 'my_devices'):
(Or press Ctrl+C to exit the script)
"@
                Write-Host $alternativeMsg -ForegroundColor Cyan -NoNewline
                $alternativeMsg | Out-File -FilePath $logFile -Append
                
                $alternativeFileName = Read-Host
                $altFileMsg = "User specified alternative file: $alternativeFileName"
                $altFileMsg | Out-File -FilePath $logFile -Append
                
                # Clean up the filename
                if ([string]::IsNullOrWhiteSpace($alternativeFileName)) {
                    $emptyMsg = "`nNo filename provided. Operation cancelled."
                    Write-Host $emptyMsg -ForegroundColor Yellow
                    $emptyMsg | Out-File -FilePath $logFile -Append
                    return
                }
                
                # Add .xlsx extension if not present
                if (-not $alternativeFileName.EndsWith(".xlsx")) {
                    $alternativeFileName = "$alternativeFileName.xlsx"
                }
                
                # Check if the specified file exists
                $alternativeFilePath = Join-Path $PSScriptRoot $alternativeFileName
                if (-not (Test-Path $alternativeFilePath)) {
                    $notFoundMsg = @"

[ERROR] File not found: $alternativeFileName
The file does not exist in the script directory: $PSScriptRoot

Operation cancelled.
"@
                    Write-Host $notFoundMsg -ForegroundColor Red
                    $notFoundMsg | Out-File -FilePath $logFile -Append
                    return
                }
                
                # Use the alternative file
                $fileName = $alternativeFileName
                $useAltMsg = "`n[OK] Using specified file: $fileName"
                Write-Host $useAltMsg -ForegroundColor Green
                $useAltMsg | Out-File -FilePath $logFile -Append
            }
            else {
                $proceedMsg = "`n[OK] Proceeding with default file: $DEFAULT_EXCEL_FILENAME"
                Write-Host $proceedMsg -ForegroundColor Green
                $proceedMsg | Out-File -FilePath $logFile -Append
            }
        }
    }
    
    $excelPath = Join-Path $PSScriptRoot $fileName
    
    $fileMessage = "`nPreparing Excel file: $excelPath"
    Write-Host $fileMessage -ForegroundColor Cyan
    $fileMessage | Out-File -FilePath $logFile -Append
    
    # Check if file exists
    if (-not (Test-Path $excelPath)) {
        $errorMessage = "ERROR: Excel file '$fileName' not found at $excelPath"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return
    }
    
    # STEP 2.1: Check if file is open
    try {
        $fileStream = [System.IO.File]::Open($excelPath, 'Open', 'Read', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        $logMessage = "[OK] File is not open and can be accessed"
        Write-Host $logMessage -ForegroundColor Green
        $logMessage | Out-File -FilePath $logFile -Append
    }
    catch {
        $errorMessage = @"

================================================================================
                                    WARNING                                       
================================================================================

The Excel file is currently OPEN in another application!

File: $excelPath

PROBLEM: The file cannot be read because it is locked by another process.
         This typically happens when the file is open in Microsoft Excel.

"@
        Write-Host $errorMessage -ForegroundColor Yellow
        $errorMessage | Out-File -FilePath $logFile -Append
        
        # Ask user if they want to close the file
        Write-Host "Would you like to close Excel and continue? (Y/N): " -ForegroundColor Cyan -NoNewline
        $response = Read-Host
        
        if ($response -match '^[Yy]') {
            $closeMsg = "`nAttempting to close Excel..."
            Write-Host $closeMsg -ForegroundColor Yellow
            $closeMsg | Out-File -FilePath $logFile -Append
            
            try {
                # Try to connect to Excel via COM to handle gracefully
                $excel = $null
                $workbookToClose = $null
                $hasUnsavedChanges = $false
                
                try {
                    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                    $comMsg = "[INFO] Connected to Excel via COM"
                    Write-Host $comMsg -ForegroundColor Cyan
                    $comMsg | Out-File -FilePath $logFile -Append
                    
                    # Find the workbook that matches our file
                    $targetFileName = [System.IO.Path]::GetFileName($excelPath)
                    foreach ($workbook in $excel.Workbooks) {
                        $workbookName = [System.IO.Path]::GetFileName($workbook.FullName)
                        if ($workbookName -eq $targetFileName) {
                            $workbookToClose = $workbook
                            $hasUnsavedChanges = -not $workbook.Saved
                            
                            $foundMsg = "[INFO] Found workbook: $workbookName (Saved: $($workbook.Saved))"
                            Write-Host $foundMsg -ForegroundColor Cyan
                            $foundMsg | Out-File -FilePath $logFile -Append
                            break
                        }
                    }
                }
                catch {
                    $comErrorMsg = "[WARNING] Could not connect to Excel via COM: $_"
                    Write-Host $comErrorMsg -ForegroundColor Yellow
                    $comErrorMsg | Out-File -FilePath $logFile -Append
                }
                
                # If we found the workbook and it has unsaved changes, ask user
                if ($workbookToClose -and $hasUnsavedChanges) {
                    Write-Host "`n[WARNING] The Excel file has UNSAVED changes!" -ForegroundColor Yellow
                    Write-Host "Do you want to save the changes before closing? (Y/N/Cancel): " -ForegroundColor Cyan -NoNewline
                    $saveResponse = Read-Host
                    
                    if ($saveResponse -match '^[Yy]') {
                        try {
                            $workbookToClose.Save()
                            $saveMsg = "[OK] Workbook saved successfully"
                            Write-Host $saveMsg -ForegroundColor Green
                            $saveMsg | Out-File -FilePath $logFile -Append
                        }
                        catch {
                            $saveErrorMsg = "[ERROR] Failed to save workbook: $_"
                            Write-Host $saveErrorMsg -ForegroundColor Red
                            $saveErrorMsg | Out-File -FilePath $logFile -Append
                            Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
                            $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                            exit
                        }
                    }
                    elseif ($saveResponse -match '^[Cc]') {
                        $cancelMsg = "`nOperation cancelled by user. Please save or close the file manually."
                        Write-Host $cancelMsg -ForegroundColor Yellow
                        $cancelMsg | Out-File -FilePath $logFile -Append
                        Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
                        $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                        exit
                    }
                    # If 'N', continue to close without saving
                }
                
                # Close via COM if we have a connection
                if ($null -ne $excel) {
                    try {
                        if ($workbookToClose) {
                            $workbookToClose.Close($false)  # False = don't save
                            $closeWbMsg = "[OK] Workbook closed via COM"
                            Write-Host $closeWbMsg -ForegroundColor Green
                            $closeWbMsg | Out-File -FilePath $logFile -Append
                        }
                        
                        # Check if there are other workbooks open
                        if ($excel.Workbooks.Count -eq 0) {
                            $excel.Quit()
                            $quitMsg = "[OK] Excel application closed (no other workbooks open)"
                            Write-Host $quitMsg -ForegroundColor Green
                            $quitMsg | Out-File -FilePath $logFile -Append
                        }
                        else {
                            $remainMsg = "[INFO] Excel remains open with other workbooks"
                            Write-Host $remainMsg -ForegroundColor Cyan
                            $remainMsg | Out-File -FilePath $logFile -Append
                        }
                        
                        # Release COM objects
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                        
                        Start-Sleep -Seconds 2
                    }
                    catch {
                        $comCloseError = "[WARNING] COM close failed: $_. Attempting forceful close."
                        Write-Host $comCloseError -ForegroundColor Yellow
                        $comCloseError | Out-File -FilePath $logFile -Append
                        
                        # Fall back to process kill
                        $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
                        if ($excelProcesses) {
                            foreach ($process in $excelProcesses) {
                                $process.CloseMainWindow() | Out-Null
                                Start-Sleep -Milliseconds 500
                                if (-not $process.HasExited) {
                                    $process | Stop-Process -Force
                                }
                            }
                            Start-Sleep -Seconds 2
                        }
                    }
                }
                else {
                    # No COM connection, fall back to process kill
                    $fallbackMsg = "[INFO] Using fallback method (process termination)"
                    Write-Host $fallbackMsg -ForegroundColor Yellow
                    $fallbackMsg | Out-File -FilePath $logFile -Append
                    
                    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
                    if ($excelProcesses) {
                        foreach ($process in $excelProcesses) {
                            $process.CloseMainWindow() | Out-Null
                            Start-Sleep -Milliseconds 500
                            if (-not $process.HasExited) {
                                $process | Stop-Process -Force
                            }
                        }
                        Start-Sleep -Seconds 2
                    }
                    else {
                        $noExcelMsg = "[WARNING] No Excel processes found, but file is still locked. May be open in another application."
                        Write-Host $noExcelMsg -ForegroundColor Yellow
                        $noExcelMsg | Out-File -FilePath $logFile -Append
                        Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
                        $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                        exit
                    }
                }
                
                # Verify file is now accessible
                try {
                    $testStream = [System.IO.File]::Open($excelPath, 'Open', 'Read', 'None')
                    $testStream.Close()
                    $testStream.Dispose()
                    
                    $successMsg = "[OK] Excel closed successfully. File is now accessible."
                    Write-Host $successMsg -ForegroundColor Green
                    $successMsg | Out-File -FilePath $logFile -Append
                }
                catch {
                    $failMsg = "[ERROR] File is still locked after closing Excel. Please close manually and try again."
                    Write-Host $failMsg -ForegroundColor Red
                    $failMsg | Out-File -FilePath $logFile -Append
                    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
                    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                    exit
                }
            }
            catch {
                $errorMsg = "[ERROR] Failed to close Excel: $_"
                Write-Host $errorMsg -ForegroundColor Red
                $errorMsg | Out-File -FilePath $logFile -Append
                Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
                $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                exit
            }
        }
        else {
            $cancelMsg = "`nOperation cancelled by user. Please close the file manually and run the script again."
            Write-Host $cancelMsg -ForegroundColor Yellow
            $cancelMsg | Out-File -FilePath $logFile -Append
            Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
            $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
            exit
        }
    }
    
    # Read Excel file
    $devices = $null
    try {
        # Check if ImportExcel module is available
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel
            $devices = Import-Excel -Path $excelPath
        }
        else {
            # Try CSV as fallback
            $csvPath = $excelPath -replace '\.xlsx$', '.csv'
            if (Test-Path $csvPath) {
                $devices = Import-Csv -Path $csvPath
            }
            else {
                throw "ImportExcel module not found and no CSV file available"
            }
        }
    }
    catch {
        $errorMessage = "ERROR: Failed to read Excel file - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return
    }
    
    # Validate required columns
    $firstDevice = $devices | Select-Object -First 1
    if (-not ($firstDevice.PSObject.Properties.Name -contains "collector_id")) {
        $errorMessage = "ERROR: Excel file must contain 'collector_id' column"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return
    }
    if (-not ($firstDevice.PSObject.Properties.Name -contains "ip_address")) {
        $errorMessage = "ERROR: Excel file must contain 'ip_address' column"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        return
    }
    
    # STEP 3: Create array of column headers
    $columnHeaders = @($firstDevice.PSObject.Properties.Name)
    $headersMessage = "`nColumn Headers Found: $($columnHeaders -join ', ')"
    Write-Host $headersMessage -ForegroundColor Cyan
    $headersMessage | Out-File -FilePath $logFile -Append
    
    # STEP 4: Create script_parameters_name array
    # Exclude: collector_id, ip_address, sample_period, and columns starting/ending with underscore
    $script_parameters_name = @()
    foreach ($header in $columnHeaders) {
        $isExcluded = ($header -eq "collector_id") -or 
        ($header -eq "ip_address") -or
        ($header -eq "sample_period") -or
        ($header -match '^_.*_$')
        
        if (-not $isExcluded) {
            $script_parameters_name += $header
        }
    }
    
    $paramsMessage = "Script Parameters Identified: $($script_parameters_name -join ', ')"
    if ($script_parameters_name.Count -eq 0) {
        $paramsMessage = "Script Parameters Identified: None (only collector_id and ip_address found)"
    }
    Write-Host $paramsMessage -ForegroundColor Cyan
    $paramsMessage | Out-File -FilePath $logFile -Append
    
    # STEP 5: Check and add _apply-result_ and _messages_ columns if not present
    $statusColumnExists = $columnHeaders -contains "_apply-result_"
    $messageColumnExists = $columnHeaders -contains "_messages_"
    
    if (-not $statusColumnExists) {
        $logMessage = "Adding '_apply-result_' column to track operation results"
        Write-Host $logMessage -ForegroundColor Yellow
        $logMessage | Out-File -FilePath $logFile -Append
        foreach ($device in $devices) {
            $device | Add-Member -NotePropertyName "_apply-result_" -NotePropertyValue "" -Force
        }
    }
    
    if (-not $messageColumnExists) {
        $logMessage = "Adding '_messages_' column to track error descriptions"
        Write-Host $logMessage -ForegroundColor Yellow
        $logMessage | Out-File -FilePath $logFile -Append
        foreach ($device in $devices) {
            $device | Add-Member -NotePropertyName "_messages_" -NotePropertyValue "" -Force
        }
    }
    
    # STEP 6: Get custom driver details including parameters
    $message = "`n=== Retrieving Custom Driver Details and Parameters ===`n"
    Write-Host $message -ForegroundColor Cyan
    $message | Out-File -FilePath $logFile -Append
    
    $customDriverDetails = $null
    $parameterMapping = @{}
    
    try {
        $apiEndpoint = "$baseURL/custom-driver/$scriptID"
        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
            "X-Api-Key"    = $apiKey
        }
        
        $logMessage = "Fetching custom driver details from: $apiEndpoint"
        Write-Host $logMessage
        $logMessage | Out-File -FilePath $logFile -Append
        
        $customDriverDetails = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get
        
        $successMessage = "[OK] Retrieved custom driver details successfully"
        Write-Host $successMessage -ForegroundColor Green
        $successMessage | Out-File -FilePath $logFile -Append
        
        # Display driver details
        $detailsMsg = @"
Custom Driver Details:
  - ID: $($customDriverDetails.id)
  - Name: $($customDriverDetails.name)
  - Type: $($customDriverDetails.type)
  - Valid: $($customDriverDetails.is_valid)
  - Requires Credentials: $($customDriverDetails.requires_credentials)
  - Has Parameters: $($customDriverDetails.code_inspection.has_parameters)
  - Has Table: $($customDriverDetails.code_inspection.has_table)
  - Has Independent Variables: $($customDriverDetails.code_inspection.has_independent_variables)
"@
        Write-Host $detailsMsg -ForegroundColor Cyan
        $detailsMsg | Out-File -FilePath $logFile -Append
        
        # STEP 6.5: Validate credentials columns if required
        if ($customDriverDetails.requires_credentials -eq $true) {
            $credValidationMsg = "`nCustom driver requires credentials - Validating username and password columns..."
            Write-Host $credValidationMsg -ForegroundColor Yellow
            $credValidationMsg | Out-File -FilePath $logFile -Append
            
            $hasUsername = $columnHeaders -contains "username"
            $hasPassword = $columnHeaders -contains "password"
            
            if (-not $hasUsername -or -not $hasPassword) {
                $missingCreds = @()
                if (-not $hasUsername) { $missingCreds += "username" }
                if (-not $hasPassword) { $missingCreds += "password" }
                
                $credErrorMsg = @"

================================================================================
                                    ERROR                                       
================================================================================

Custom driver '$($customDriverDetails.name)' requires credentials!

PROBLEM: The custom driver has 'requires_credentials' set to true, but the
         Excel file is missing required credential columns.

Missing columns: $($missingCreds -join ', ')

SOLUTION: Add the following columns to your Excel file:
         - username: Username for device authentication
         - password: Password for device authentication

Current Excel columns: $($columnHeaders -join ', ')

Script execution stopped.
================================================================================
"@
                Write-Host $credErrorMsg -ForegroundColor Red
                $credErrorMsg | Out-File -FilePath $logFile -Append
                throw "Missing required credential columns: $($missingCreds -join ', ')"
            }
            
            $credSuccessMsg = "[OK] Credentials columns validated: username and password are present"
            Write-Host $credSuccessMsg -ForegroundColor Green
            $credSuccessMsg | Out-File -FilePath $logFile -Append
        }
        else {
            $noCredsMsg = "[INFO] Custom driver does not require credentials"
            Write-Host $noCredsMsg -ForegroundColor Gray
            $noCredsMsg | Out-File -FilePath $logFile -Append
        }
        
        # STEP 7: Map script_parameters_name to parameter IDs
        if ($customDriverDetails.parameters -and $customDriverDetails.parameters.Count -gt 0) {
            $paramMessage = "`nCustom Driver Parameters Found: $($customDriverDetails.parameters.Count)"
            Write-Host $paramMessage -ForegroundColor Cyan
            $paramMessage | Out-File -FilePath $logFile -Append
            
            # Arrays to track mismatches
            $excelParamsNotInDriver = @()
            $driverParamsNotInExcel = @()
            
            # Create mapping (excluding username and password which are credentials, not parameters)
            foreach ($excelParam in $script_parameters_name) {
                # Skip credential fields - these are not script parameters
                if ($excelParam -eq "username" -or $excelParam -eq "password") {
                    continue
                }
                
                # Parse Excel column name to extract parameter name and type
                # Format: "parameter_name (TYPE)"
                $paramName = $excelParam
                $paramType = $null
                
                if ($excelParam -match '^(.+?)\s*\(([^)]+)\)\s*$') {
                    $paramName = $matches[1].Trim()
                    $paramType = $matches[2].Trim()
                }
                
                # Find matching parameter by name
                $matchingParam = $customDriverDetails.parameters | Where-Object { $_.name -eq $paramName }
                if ($matchingParam) {
                    # Verify type matches if it was specified in Excel column
                    if ($paramType -and $matchingParam.value_type -ne $paramType) {
                        $typeMismatchMsg = "  [WARNING] Excel column '$excelParam' has type '$paramType' but driver parameter '$paramName' has type '$($matchingParam.value_type)'"
                        Write-Host $typeMismatchMsg -ForegroundColor Yellow
                        $typeMismatchMsg | Out-File -FilePath $logFile -Append
                        $excelParamsNotInDriver += $excelParam
                    }
                    else {
                        # Mapping successful - use original Excel column name as key
                        $parameterMapping[$excelParam] = @{
                            id          = $matchingParam.id
                            name        = $paramName
                            value_type  = $matchingParam.value_type
                            description = $matchingParam.description
                        }
                        $mappingMsg = "  [OK] Mapped Excel column '$excelParam' to Parameter ID: $($matchingParam.id) (Name: $paramName, Type: $($matchingParam.value_type))"
                        Write-Host $mappingMsg -ForegroundColor Green
                        $mappingMsg | Out-File -FilePath $logFile -Append
                    }
                }
                else {
                    # Excel has a parameter that's not in the driver
                    $excelParamsNotInDriver += $excelParam
                }
            }
            
            # Check for driver parameters not in Excel
            # Extract parameter names from Excel columns (removing type suffix if present)
            $excelParamNames = @()
            foreach ($excelParam in $script_parameters_name) {
                # Skip credentials
                if ($excelParam -eq "username" -or $excelParam -eq "password") {
                    continue
                }
                # Extract name from "name (TYPE)" format
                if ($excelParam -match '^(.+?)\s*\(([^)]+)\)\s*$') {
                    $excelParamNames += $matches[1].Trim()
                }
                else {
                    $excelParamNames += $excelParam
                }
            }
            
            foreach ($driverParam in $customDriverDetails.parameters) {
                if ($driverParam.name -notin $excelParamNames) {
                    $driverParamsNotInExcel += $driverParam.name
                }
            }
            
            # CRITICAL VALIDATION: Excel and Driver parameters must match exactly
            if ($excelParamsNotInDriver.Count -gt 0 -or $driverParamsNotInExcel.Count -gt 0) {
                $mismatchError = @"

================================================================================
                           PARAMETER MISMATCH ERROR                            
================================================================================

ERROR: Excel file parameters do NOT match the custom driver parameters!

"@
                if ($excelParamsNotInDriver.Count -gt 0) {
                    $mismatchError += @"
Excel columns NOT found in custom driver parameters:
$(($excelParamsNotInDriver | ForEach-Object { "  - $_" }) -join "`n")

"@
                }
                
                if ($driverParamsNotInExcel.Count -gt 0) {
                    $mismatchError += @"
Custom driver parameters NOT found in Excel columns:
$(($driverParamsNotInExcel | ForEach-Object { "  - $_" }) -join "`n")

"@
                }
                
                $mismatchError += @"

Expected custom driver parameters:
$(($customDriverDetails.parameters | ForEach-Object { "  - $($_.name)" }) -join "`n")

Excel parameter columns found (excluding credentials):
$(($script_parameters_name | Where-Object { $_ -ne "username" -and $_ -ne "password" } | ForEach-Object { "  - $_" }) -join "`n")

SOLUTION:
1. Re-create the Excel file using the -operation create-excel command
   with the correct -script_name to ensure parameters match
2. OR manually add/remove columns in the Excel file to match exactly

Script execution stopped.
================================================================================
"@
                Write-Host $mismatchError -ForegroundColor Red
                $mismatchError | Out-File -FilePath $logFile -Append
                
                # Open Excel file to help user fix the issue
                if ([string]::IsNullOrEmpty($fileName)) {
                    $fileName = $DEFAULT_EXCEL_FILENAME
                }
                if (-not $fileName.EndsWith(".xlsx")) {
                    $fileName = "$fileName.xlsx"
                }
                $excelPath = Join-Path $PSScriptRoot $fileName
                
                if (Test-Path $excelPath) {
                    $openMsg = "`nOpening Excel file for review..."
                    Write-Host $openMsg -ForegroundColor Yellow
                    $openMsg | Out-File -FilePath $logFile -Append
                    Start-Process $excelPath
                }
                
                throw "Parameter mismatch between Excel file and custom driver"
            }
            
            $validationMsg = "`n[OK] All Excel parameters match custom driver parameters perfectly"
            Write-Host $validationMsg -ForegroundColor Green
            $validationMsg | Out-File -FilePath $logFile -Append
        }
        else {
            $noParamsMsg = "`nNo parameters defined for this custom driver"
            Write-Host $noParamsMsg -ForegroundColor Yellow
            $noParamsMsg | Out-File -FilePath $logFile -Append
        }
    }
    catch {
        $errorMessage = "ERROR: Failed to retrieve custom driver details - $_"
        Write-Host $errorMessage -ForegroundColor Red
        $errorMessage | Out-File -FilePath $logFile -Append
        
        # Open Excel file to help user review
        if ([string]::IsNullOrEmpty($fileName)) {
            $fileName = $DEFAULT_EXCEL_FILENAME
        }
        if (-not $fileName.EndsWith(".xlsx")) {
            $fileName = "$fileName.xlsx"
        }
        $excelPath = Join-Path $PSScriptRoot $fileName
        
        if (Test-Path $excelPath) {
            $openMsg = "`nOpening Excel file for review..."
            Write-Host $openMsg -ForegroundColor Yellow
            $openMsg | Out-File -FilePath $logFile -Append
            Start-Process $excelPath
        }
        
        return
    }
    
    # PRE-ANALYSIS: Scan for DeleteAssociation and UpdateParameters operations
    $preAnalysisMsg = "`n=== Pre-Analysis: Scanning for DeleteAssociation/UpdateParameters Operations ===`n"
    Write-Host $preAnalysisMsg -ForegroundColor Magenta
    $preAnalysisMsg | Out-File -FilePath $logFile -Append
    
    # Check if _operation_ column exists
    $hasOperationColumn = $columnHeaders -contains "_operation_"
    $collectorsWithSpecialOps = @{}
    $allExistingAssociations = @()
    
    if ($hasOperationColumn) {
        $scanMsg = "[INFO] Scanning all rows for DeleteAssociation or UpdateParameters operations..."
        Write-Host $scanMsg -ForegroundColor Cyan
        $scanMsg | Out-File -FilePath $logFile -Append
        
        # Scan all devices for special operations
        $collectorsToQuery = @{}
        foreach ($device in $devices) {
            $operation = $device."_operation_"
            if ($operation -eq "DeleteAssociation" -or $operation -eq "UpdateParameters") {
                $collectorId = $device.collector_id
                if (-not [string]::IsNullOrWhiteSpace($collectorId)) {
                    if (-not $collectorsToQuery.ContainsKey($collectorId)) {
                        $collectorsToQuery[$collectorId] = @{
                            "DeleteAssociation" = 0
                            "UpdateParameters"  = 0
                        }
                    }
                    $collectorsToQuery[$collectorId][$operation]++
                }
            }
        }
        
        if ($collectorsToQuery.Count -gt 0) {
            $foundMsg = @"

[FOUND] Detected special operations in the spreadsheet:
  - Collectors with DeleteAssociation or UpdateParameters: $($collectorsToQuery.Count)

Collector breakdown:
"@
            Write-Host $foundMsg -ForegroundColor Yellow
            $foundMsg | Out-File -FilePath $logFile -Append
            
            foreach ($collId in $collectorsToQuery.Keys) {
                $disassocCount = $collectorsToQuery[$collId]["DeleteAssociation"]
                $updateCount = $collectorsToQuery[$collId]["UpdateParameters"]
                $breakdownMsg = "  - Collector ID $collId`: DeleteAssociation=$disassocCount, UpdateParameters=$updateCount"
                Write-Host $breakdownMsg -ForegroundColor Cyan
                $breakdownMsg | Out-File -FilePath $logFile -Append
            }
            
            # Fetch existing associations for each collector
            $fetchMsg = "`n[INFO] Fetching existing associations from Domotz API..."
            Write-Host $fetchMsg -ForegroundColor Cyan
            $fetchMsg | Out-File -FilePath $logFile -Append
            
            foreach ($collId in $collectorsToQuery.Keys) {
                try {
                    $assocEndpoint = "$baseURL/custom-driver/agent/$collId/association"
                    $assocHeaders = @{
                        "Accept"       = "application/json"
                        "Content-Type" = "application/json"
                        "X-Api-Key"    = $apiKey
                    }
                    
                    $assocMsg = "  - Fetching associations from Collector ID $collId..."
                    Write-Host $assocMsg -ForegroundColor Cyan
                    $assocMsg | Out-File -FilePath $logFile -Append
                    
                    $associations = Invoke-RestMethod -Uri $assocEndpoint -Headers $assocHeaders -Method Get
                    
                    # Filter to only associations matching the current script ID
                    $matchingAssociations = $associations | Where-Object { $_.custom_driver_id -eq $scriptID }
                    
                    if ($matchingAssociations) {
                        $matchCount = @($matchingAssociations).Count
                        $matchMsg = "    [OK] Found $matchCount association(s) for script '$scriptName' (ID: $scriptID)"
                        Write-Host $matchMsg -ForegroundColor Green
                        $matchMsg | Out-File -FilePath $logFile -Append
                        
                        # Store associations for this collector with collector_id added
                        foreach ($assoc in $matchingAssociations) {
                            # Add collector_id to the association object for later reference
                            $assoc | Add-Member -NotePropertyName "collector_id" -NotePropertyValue $collId -Force
                        }
                        $collectorsWithSpecialOps[$collId] = $matchingAssociations
                        $allExistingAssociations += $matchingAssociations
                    }
                    else {
                        $noMatchMsg = "    [INFO] No associations found for script '$scriptName' (ID: $scriptID)"
                        Write-Host $noMatchMsg -ForegroundColor Gray
                        $noMatchMsg | Out-File -FilePath $logFile -Append
                    }
                }
                catch {
                    $errorMsg = "    [ERROR] Failed to fetch associations from Collector ID $collId`: $_"
                    Write-Host $errorMsg -ForegroundColor Red
                    $errorMsg | Out-File -FilePath $logFile -Append
                }
            }
            
            # Display summary of found associations (detailed output only in debug mode)
            if ($allExistingAssociations.Count -gt 0) {
                if ($debug) {
                    $summaryMsg = @"

================================================================================
                  EXISTING ASSOCIATIONS FOUND                                   
================================================================================

Total associations for script '$scriptName' (ID: $scriptID): $($allExistingAssociations.Count)

"@
                    Write-Host $summaryMsg -ForegroundColor Green
                    $summaryMsg | Out-File -FilePath $logFile -Append
                    
                    # Display each association
                    $index = 1
                    foreach ($assoc in $allExistingAssociations) {
                        $assocDetails = @"
Association #${index}:
  - Collector ID: $($assoc.collector_id)
  - Association ID: $($assoc.id)
  - Device ID: $($assoc.device_id)
  - Custom Driver ID: $($assoc.custom_driver_id)
  - Status: $($assoc.status)
  - Sample Period: $($assoc.sample_period)
  - Last Inspection: $($assoc.last_inspection_time)
  - Used Variables: $($assoc.used_variables)
"@
                        Write-Host $assocDetails -ForegroundColor Cyan
                        $assocDetails | Out-File -FilePath $logFile -Append
                        
                        if ($assoc.parameters -and $assoc.parameters.Count -gt 0) {
                            $paramListMsg = "  - Parameters ($($assoc.parameters.Count)):"
                            Write-Host $paramListMsg -ForegroundColor Cyan
                            $paramListMsg | Out-File -FilePath $logFile -Append
                            
                            foreach ($param in $assoc.parameters) {
                                $paramDetail = "      * $($param.name) ($($param.value_type)): $($param.value)"
                                Write-Host $paramDetail -ForegroundColor Gray
                                $paramDetail | Out-File -FilePath $logFile -Append
                            }
                        }
                        Write-Host ""
                        "" | Out-File -FilePath $logFile -Append
                        $index++
                    }
                    
                    $endSummaryMsg = "================================================================================"
                    Write-Host $endSummaryMsg -ForegroundColor Green
                    $endSummaryMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    $briefSummaryMsg = "`n[INFO] Found $($allExistingAssociations.Count) existing association(s) for script '$scriptName'. Use -debug flag to see details."
                    Write-Host $briefSummaryMsg -ForegroundColor Cyan
                    $briefSummaryMsg | Out-File -FilePath $logFile -Append
                }
            }
            else {
                $noAssocMsg = "`n[INFO] No existing associations found for script '$scriptName' in the specified collectors"
                Write-Host $noAssocMsg -ForegroundColor Yellow
                $noAssocMsg | Out-File -FilePath $logFile -Append
            }
        }
        else {
            $noSpecialOpsMsg = "[INFO] No DeleteAssociation or UpdateParameters operations found in the spreadsheet"
            Write-Host $noSpecialOpsMsg -ForegroundColor Gray
            $noSpecialOpsMsg | Out-File -FilePath $logFile -Append
        }
    }
    else {
        $noColumnMsg = "[INFO] _operation_ column not found in spreadsheet - skipping pre-analysis"
        Write-Host $noColumnMsg -ForegroundColor Gray
        $noColumnMsg | Out-File -FilePath $logFile -Append
    }
    
    # Initialize counters
    $script:totalAttempts = 0
    $script:successCount = 0
    $script:failureCount = 0
    $script:skippedCount = 0
    $script:successDetails = @()
    $script:failureDetails = @()
    $script:skippedDetails = @()
    
    # Cache for device lists to avoid repeated API calls
    $deviceListCache = @{}
    
    # Start bulk apply operation
    $startBulkMsg = "`n=== Starting Bulk Apply Operation ===`n"
    Write-Host $startBulkMsg -ForegroundColor Magenta
    $startBulkMsg | Out-File -FilePath $logFile -Append
    
    $processingMsg = "[INFO] Processing $($devices.Count) row(s) from spreadsheet..."
    Write-Host $processingMsg -ForegroundColor Cyan
    $processingMsg | Out-File -FilePath $logFile -Append
    
    # Process each device
    $rowNumber = 1
    foreach ($device in $devices) {
        $script:totalAttempts++
        
        # CHECK: Only process rows with _operation_ specified
        if ($hasOperationColumn) {
            $operation = $device."_operation_"
            if ([string]::IsNullOrWhiteSpace($operation)) {
                $rowNumber++
                $skipOperationMsg = "[Row #$($rowNumber-1)] SKIPPED - No operation specified in _operation_ column (Collector: $($device.collector_id), IP: $($device.ip_address))"
                Write-Host $skipOperationMsg -ForegroundColor Gray
                $skipOperationMsg | Out-File -FilePath $logFile -Append
                continue
            }
        }
        
        # EARLY CHECK: Validate required parameters before processing row
        # This prevents verbose output for rows that will be skipped anyway
        $missingRequiredParams = @()
        
        # Check all script parameters
        foreach ($paramName in $script_parameters_name) {
            $paramValue = if ($device.PSObject.Properties.Name -contains $paramName) { $device.$paramName } else { "" }
            if ([string]::IsNullOrWhiteSpace($paramValue)) {
                $missingRequiredParams += $paramName
            }
        }
        
        # Check credentials if required
        if ($customDriverDetails.requires_credentials -eq $true) {
            if ([string]::IsNullOrWhiteSpace($device.username)) {
                $missingRequiredParams += "username"
            }
            if ([string]::IsNullOrWhiteSpace($device.password)) {
                $missingRequiredParams += "password"
            }
        }
        
        # Check sample_period
        if ($device.PSObject.Properties.Name -contains "sample_period") {
            if ([string]::IsNullOrWhiteSpace($device.sample_period)) {
                $missingRequiredParams += "sample_period"
            }
        }
        else {
            $missingRequiredParams += "sample_period"
        }
        
        # If any required parameters are missing, skip immediately without verbose output
        if ($missingRequiredParams.Count -gt 0) {
            $rowNumber++
            $skipMsg = "[Row #$($rowNumber-1)] SKIPPED - Missing: $($missingRequiredParams -join ', ') (Collector: $($device.collector_id), IP: $($device.ip_address))"
            Write-Host $skipMsg -ForegroundColor Yellow
            $skipMsg | Out-File -FilePath $logFile -Append
            
            $device."_apply-result_" = "Skipped"
            $device._messages_ = "Missing required parameters: $($missingRequiredParams -join ', ')"
            
            $script:skippedCount++
            $script:skippedDetails += "Row #$($rowNumber-1): Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - Missing: $($missingRequiredParams -join ', ')"
            continue
        }
        
        # Log the Excel row content for rows that will be processed
        $separator = "-" * 80
        $rowMessage = "`n$separator`nProcessing Excel Row #$rowNumber`n$separator"
        Write-Host $rowMessage -ForegroundColor Yellow
        $rowMessage | Out-File -FilePath $logFile -Append
        
        # Output all columns from the Excel row for troubleshooting
        $rowDetails = "Excel Row Content:"
        Write-Host $rowDetails
        $rowDetails | Out-File -FilePath $logFile -Append
        
        foreach ($property in $device.PSObject.Properties) {
            $propertyLine = "  $($property.Name): $($property.Value)"
            Write-Host $propertyLine -ForegroundColor White
            $propertyLine | Out-File -FilePath $logFile -Append
        }
        
        $rowNumber++
        
        # STEP 1: Skip if collector_id or ip_address is empty
        if ([string]::IsNullOrWhiteSpace($device.collector_id) -or [string]::IsNullOrWhiteSpace($device.ip_address)) {
            $warningMessage = "`n[WARNING] Skipping row with empty collector_id or ip_address"
            Write-Host $warningMessage -ForegroundColor Yellow
            $warningMessage | Out-File -FilePath $logFile -Append
            $device."_apply-result_" = "Error"
            $device._messages_ = "Empty collector_id or ip_address"
            continue
        }
        
        # STEP 2: Get device ID - use existing _device_id_ if available, otherwise retrieve from API
        $deviceID = $null
        
        # Check if _device_id_ is already provided in the Excel row
        if ($device.PSObject.Properties.Name -contains "_device_id_" -and 
            -not [string]::IsNullOrWhiteSpace($device."_device_id_")) {
            $deviceID = $device."_device_id_"
            $cachedMsg = "`nUsing existing Device ID from Excel: $deviceID (Skipping API call)"
            Write-Host $cachedMsg -ForegroundColor Cyan
            $cachedMsg | Out-File -FilePath $logFile -Append
        }
        else {
            # STEP 3: Get device list if not in cache
            if (-not $deviceListCache.ContainsKey($device.collector_id)) {
                $deviceListCache[$device.collector_id] = Get-DeviceList -collectorID $device.collector_id
            }
            
            # STEP 4: Get device ID from IP via API
            $retrievalMessage = "`nAttempting to retrieve Domotz Device ID from IP: $($device.ip_address)"
            Write-Host $retrievalMessage
            $retrievalMessage | Out-File -FilePath $logFile -Append
            
            $deviceID = Get-DeviceIDFromIP -deviceIP $device.ip_address -collectorID $device.collector_id -deviceList $deviceListCache[$device.collector_id]
            
            if (-not $deviceID) {
                $errorMsg = "`n>>> ERROR: Could not retrieve Domotz Device ID for IP: $($device.ip_address) <<<"
                Write-Host $errorMsg -ForegroundColor Red
                $errorMsg | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "Error"
                $device._messages_ = "Device ID not found for IP $($device.ip_address)"
                
                $script:failureCount++
                $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - Device ID not found"
                continue
            }
            
            # STEP 5: Device ID retrieved successfully
            $idMessage = "`n>>> RESULT: Domotz Device ID = $deviceID <<<"
            Write-Host $idMessage -ForegroundColor Green
            $idMessage | Out-File -FilePath $logFile -Append
        }
        
        # STEP 6: Build parameters array for API call and validate LIST types
        $parametersArray = @()
        $paramValidationErrors = @()
        
        foreach ($key in $parameterMapping.Keys) {
            $paramInfo = $parameterMapping[$key]
            $paramValue = $device.$key
            
            if (-not [string]::IsNullOrWhiteSpace($paramValue)) {
                # For logging: mask value if SECRET_TEXT
                $displayValue = if ($paramInfo.value_type -eq "SECRET_TEXT") { "********" } else { $paramValue }
                
                # Check if parameter type is LIST
                if ($paramInfo.value_type -eq "LIST") {
                    # Validate that the value is in array format: ["value1", "value2"] or []
                    $trimmedValue = $paramValue.Trim()
                    
                    # Check if it starts with [ and ends with ]
                    if ($trimmedValue -match '^\s*\[.*\]\s*$') {
                        # Check if it's an empty array
                        if ($trimmedValue -match '^\s*\[\s*\]\s*$') {
                            # Empty array - create explicit empty array with proper type
                            $parametersArray += [PSCustomObject]@{
                                custom_driver_parameter_id = $paramInfo.id
                                value                      = [object[]]@()
                            }
                            
                            $validMsg = "  [OK] Parameter '$key' validated as LIST (empty array): $trimmedValue"
                            Write-Host $validMsg -ForegroundColor Green
                            $validMsg | Out-File -FilePath $logFile -Append
                        }
                        else {
                            # Try to parse as JSON array
                            try {
                                $arrayValue = $trimmedValue | ConvertFrom-Json
                                
                                # Verify it's actually an array
                                if ($arrayValue -is [System.Array] -or $arrayValue -is [System.Collections.ArrayList]) {
                                    # Convert to a proper object array for correct JSON serialization
                                    # Cast explicitly to [object[]] to avoid PowerShell adding metadata
                                    $plainArray = [object[]]($arrayValue | ForEach-Object { $_ })
                                    $parametersArray += [PSCustomObject]@{
                                        custom_driver_parameter_id = $paramInfo.id
                                        value                      = $plainArray
                                    }
                                    
                                    $validMsg = "  [OK] Parameter '$key' validated as LIST: $trimmedValue"
                                    Write-Host $validMsg -ForegroundColor Green
                                    $validMsg | Out-File -FilePath $logFile -Append
                                }
                                else {
                                    $paramValidationErrors += "Parameter '$key' (LIST): Value must be a JSON array like [`"value1`", `"value2`"], got: $displayValue"
                                }
                            }
                            catch {
                                $paramValidationErrors += "Parameter '$key' (LIST): Invalid JSON array format. Expected [`"value1`", `"value2`"], got: $displayValue. Error: $_"
                            }
                        }
                    }
                    else {
                        $paramValidationErrors += "Parameter '$key' (LIST): Value must be in array format [`"value1`", `"value2`"], got: $displayValue"
                    }
                }
                elseif ($paramInfo.value_type -eq "SECRET_TEXT") {
                    # SECRET_TEXT parameter - use as string, already masked for display above
                    $parametersArray += [PSCustomObject]@{
                        custom_driver_parameter_id = $paramInfo.id
                        value                      = $paramValue
                    }
                    
                    $validMsg = "  [OK] Parameter '$key' validated as SECRET_TEXT: ********"
                    Write-Host $validMsg -ForegroundColor Green
                    $validMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    # Non-LIST, non-SECRET_TEXT parameter - use as string
                    $parametersArray += [PSCustomObject]@{
                        custom_driver_parameter_id = $paramInfo.id
                        value                      = $paramValue
                    }
                }
            }
        }
        
        # If there are validation errors, mark row as error and skip
        if ($paramValidationErrors.Count -gt 0) {
            $errorSummary = $paramValidationErrors -join '; '
            $validationErrorMsg = "`n[VALIDATION ERROR] Parameter type mismatch:"
            Write-Host $validationErrorMsg -ForegroundColor Red
            $validationErrorMsg | Out-File -FilePath $logFile -Append
            
            foreach ($err in $paramValidationErrors) {
                $errMsg = "  - $err"
                Write-Host $errMsg -ForegroundColor Red
                $errMsg | Out-File -FilePath $logFile -Append
            }
            
            $device."_apply-result_" = "Error"
            $device._messages_ = "Validation error: $errorSummary"
            
            $script:failureCount++
            $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - $errorSummary"
            continue
        }
        
        # STEP 6: Get sample_period from the row
        $samplePeriod = 300  # Default value
        if ($device.PSObject.Properties.Name -contains "sample_period") {
            if (-not [string]::IsNullOrWhiteSpace($device.sample_period)) {
                $samplePeriod = [int]$device.sample_period
            }
        }
        
        # STEP 7: Display API call information for troubleshooting (only if debug mode)
        if ($debug) {
            # Mask API key for security
            $maskedApiKey = "****" + $apiKey.Substring([Math]::Max(0, $apiKey.Length - 4))
            
            $apiCallInfo = @"

================================================================================
                        API CALL INFORMATION (TROUBLESHOOTING)                        
================================================================================

API Endpoint:
POST $baseURL/custom-driver/$scriptID/agent/$($device.collector_id)/device/$deviceID/association

Headers:
  Content-Type: application/json
  X-Api-Key: $maskedApiKey

Request Body:
{
  "parameters": [
"@
            Write-Host $apiCallInfo -ForegroundColor Magenta
            $apiCallInfo | Out-File -FilePath $logFile -Append
            
            # Display parameters
            for ($i = 0; $i -lt $parametersArray.Count; $i++) {
                $param = $parametersArray[$i]
                $comma = if ($i -lt ($parametersArray.Count - 1)) { "," } else { "" }
                
                # Find parameter type from parameterMapping to check if it's SECRET_TEXT
                $isSecretParam = $false
                foreach ($key in $parameterMapping.Keys) {
                    $paramInfo = $parameterMapping[$key]
                    if ($paramInfo.id -eq $param.custom_driver_parameter_id -and $paramInfo.value_type -eq "SECRET_TEXT") {
                        $isSecretParam = $true
                        break
                    }
                }
                
                # Format value based on type
                if ($param.value -is [System.Array] -or $param.value -is [System.Collections.ArrayList]) {
                    # LIST type - display as JSON array
                    if ($param.value.Count -eq 0) {
                        # Empty array - explicitly show as []
                        $valueJson = "[]"
                    }
                    else {
                        $valueJson = ($param.value | ConvertTo-Json -Compress)
                    }
                    $paramLine = "    { `"custom_driver_parameter_id`": $($param.custom_driver_parameter_id), `"value`": $valueJson }$comma"
                }
                else {
                    # String type - mask if SECRET_TEXT, otherwise display with quotes
                    if ($isSecretParam) {
                        $paramLine = "    { `"custom_driver_parameter_id`": $($param.custom_driver_parameter_id), `"value`": `"********`" }$comma"
                    }
                    else {
                        $paramLine = "    { `"custom_driver_parameter_id`": $($param.custom_driver_parameter_id), `"value`": `"$($param.value)`" }$comma"
                    }
                }
                
                Write-Host $paramLine -ForegroundColor Cyan
                $paramLine | Out-File -FilePath $logFile -Append
            }
            
            $bodyEnd = @"
  ],
  "sample_period": $samplePeriod
"@
            
            # Add credentials to display if required
            if ($customDriverDetails.requires_credentials -eq $true) {
                $username = if ($device.PSObject.Properties.Name -contains "username") { $device.username } else { "" }
                $password = if ($device.PSObject.Properties.Name -contains "password") { $device.password } else { "" }
                $bodyEnd += @"
,
  "credentials": {
    "username": "$username",
    "password": "********"
  }
"@
            }
            
            $bodyEnd += @"
}

Row Data Summary:
  - Collector ID (AGENT_ID): $($device.collector_id)
  - Device IP: $($device.ip_address)
  - Device ID (DEVICE_ID): $deviceID
  - Custom Script ID (CUSTOM_SCRIPT_ID): $scriptID
  - Sample Period: $samplePeriod
  - Parameters Count: $($parametersArray.Count)
"@
            
            if ($customDriverDetails.requires_credentials -eq $true) {
                $bodyEnd += "`n  - Requires Credentials: YES (username and password included)"
            }
            else {
                $bodyEnd += "`n  - Requires Credentials: NO"
            }
            
            $bodyEnd += @"

================================================================================
"@
            Write-Host $bodyEnd -ForegroundColor Magenta
            $bodyEnd | Out-File -FilePath $logFile -Append
        }
        
        # STEP 8: Make the API call based on operation type
        # Get operation type (default to "Associate" if not specified for backward compatibility)
        $operationType = if ($hasOperationColumn -and $device.PSObject.Properties.Name -contains "_operation_") { 
            $device."_operation_" 
        }
        else { 
            "Associate" 
        }
        
        if ($operationType -eq "Associate") {
            # STEP 8A: Associate custom driver to device
            try {
                $apiEndpoint = "$baseURL/custom-driver/$scriptID/agent/$($device.collector_id)/device/$deviceID/association"
                $headers = @{
                    "Accept"       = "application/json"
                    "Content-Type" = "application/json"
                    "X-Api-Key"    = $apiKey
                }
                
                # Build request body using PSCustomObject for proper JSON serialization
                if ($customDriverDetails.requires_credentials -eq $true) {
                    $username = if ($device.PSObject.Properties.Name -contains "username") { $device.username } else { "" }
                    $password = if ($device.PSObject.Properties.Name -contains "password") { $device.password } else { "" }
                    
                    $requestBodyObj = [PSCustomObject]@{
                        parameters    = $parametersArray
                        sample_period = $samplePeriod
                        credentials   = [PSCustomObject]@{
                            username = $username
                            password = $password
                        }
                    }
                    
                    $credMsg = "[INFO] Including credentials in API request (username: $username)"
                    Write-Host $credMsg -ForegroundColor Gray
                    $credMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    $requestBodyObj = [PSCustomObject]@{
                        parameters    = $parametersArray
                        sample_period = $samplePeriod
                    }
                }
                
                $requestBody = $requestBodyObj | ConvertTo-Json -Depth 100 -Compress:$false
                
                # Debug: Log the actual JSON being sent (only if -debug parameter is set)
                if ($debug) {
                    # Create a masked copy for debug output (mask SECRET_TEXT parameters)
                    $maskedParametersArray = @()
                    foreach ($param in $parametersArray) {
                        # Find if this parameter is SECRET_TEXT
                        $isSecretParam = $false
                        foreach ($key in $parameterMapping.Keys) {
                            $paramInfo = $parameterMapping[$key]
                            if ($paramInfo.id -eq $param.custom_driver_parameter_id -and $paramInfo.value_type -eq "SECRET_TEXT") {
                                $isSecretParam = $true
                                break
                            }
                        }
                        
                        if ($isSecretParam) {
                            # Mask the secret value
                            $maskedParametersArray += [PSCustomObject]@{
                                custom_driver_parameter_id = $param.custom_driver_parameter_id
                                value                      = "********"
                            }
                        }
                        else {
                            # Keep original value
                            $maskedParametersArray += $param
                        }
                    }
                    
                    # Build masked request body for debug display
                    if ($customDriverDetails.requires_credentials -eq $true) {
                        $maskedRequestBodyObj = [PSCustomObject]@{
                            parameters    = $maskedParametersArray
                            sample_period = $samplePeriod
                            credentials   = [PSCustomObject]@{
                                username = $username
                                password = "********"
                            }
                        }
                    }
                    else {
                        $maskedRequestBodyObj = [PSCustomObject]@{
                            parameters    = $maskedParametersArray
                            sample_period = $samplePeriod
                        }
                    }
                    
                    $maskedRequestBody = $maskedRequestBodyObj | ConvertTo-Json -Depth 100 -Compress:$false
                    $debugMsg = "`n[DEBUG] Actual JSON Request Body being sent to API (SECRET_TEXT parameters masked):`n$maskedRequestBody"
                    Write-Host $debugMsg -ForegroundColor Gray
                    $debugMsg | Out-File -FilePath $logFile -Append
                }
                
                $callMessage = "`n[API CALL] Associating custom driver..."
                Write-Host $callMessage -ForegroundColor Yellow
                $callMessage | Out-File -FilePath $logFile -Append
                
                $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Post -Body $requestBody
                
                $successMessage = "[SUCCESS] Custom driver associated successfully"
                Write-Host $successMessage -ForegroundColor Green
                $successMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "OK"
                $device._messages_ = "Successfully associated"
                
                $script:successCount++
                $script:successDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address), Device ID: $deviceID"
            }
            catch {
                $errorMessage = "[ERROR] Failed to associate custom driver: $_"
                Write-Host $errorMessage -ForegroundColor Red
                $errorMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "Error"
                $device._messages_ = "API call failed: $_"
                
                $script:failureCount++
                $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - API Error: $_"
            }
        }
        elseif ($operationType -eq "DeleteAssociation") {
            # STEP 8B: Delete association from device
            try {
                # Find the association ID for this device from the pre-analysis data
                $associationId = $null
                $matchingAssoc = $allExistingAssociations | Where-Object { 
                    $_.device_id -eq $deviceID -and $_.collector_id -eq $device.collector_id 
                }
                
                if ($matchingAssoc) {
                    $associationId = $matchingAssoc.id
                    
                    $foundAssocMsg = "[INFO] Found existing association ID: $associationId for Device ID: $deviceID"
                    Write-Host $foundAssocMsg -ForegroundColor Cyan
                    $foundAssocMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    # Association not found in pre-analysis, throw error
                    throw "No existing association found for this device. Cannot delete non-existent association."
                }
                
                # Make DELETE API call
                $apiEndpoint = "$baseURL/custom-driver/$scriptID/association/$associationId"
                $headers = @{
                    "Accept"       = "application/json"
                    "Content-Type" = "application/json"
                    "X-Api-Key"    = $apiKey
                }
                
                $callMessage = "`n[API CALL] Deleting association (ID: $associationId)..."
                Write-Host $callMessage -ForegroundColor Yellow
                $callMessage | Out-File -FilePath $logFile -Append
                
                $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Delete
                
                $successMessage = "[SUCCESS] Association deleted successfully"
                Write-Host $successMessage -ForegroundColor Green
                $successMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "OK"
                $device._messages_ = "Association deleted successfully"
                
                $script:successCount++
                $script:successDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address), Device ID: $deviceID - Deleted Association ID: $associationId"
            }
            catch {
                $errorMessage = "[ERROR] Failed to delete association: $_"
                Write-Host $errorMessage -ForegroundColor Red
                $errorMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "Error"
                $device._messages_ = "Delete operation failed: $_"
                
                $script:failureCount++
                $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - Delete Error: $_"
            }
        }
        elseif ($operationType -eq "UpdateParameters") {
            # STEP 8C: Update existing association parameters
            try {
                # Find the association ID for this device from the pre-analysis data
                $associationId = $null
                $matchingAssoc = $allExistingAssociations | Where-Object { 
                    $_.device_id -eq $deviceID -and $_.collector_id -eq $device.collector_id 
                }
                
                if ($matchingAssoc) {
                    $associationId = $matchingAssoc.id
                    
                    $foundAssocMsg = "[INFO] Found existing association ID: $associationId for Device ID: $deviceID"
                    Write-Host $foundAssocMsg -ForegroundColor Cyan
                    $foundAssocMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    # Association not found in pre-analysis, throw error
                    throw "No existing association found for this device. Cannot update non-existent association."
                }
                
                # Build request body using PSCustomObject for proper JSON serialization (same as Associate)
                $apiEndpoint = "$baseURL/custom-driver/$scriptID/association/$associationId"
                $headers = @{
                    "Accept"       = "application/json"
                    "Content-Type" = "application/json"
                    "X-Api-Key"    = $apiKey
                }
                
                if ($customDriverDetails.requires_credentials -eq $true) {
                    $username = if ($device.PSObject.Properties.Name -contains "username") { $device.username } else { "" }
                    $password = if ($device.PSObject.Properties.Name -contains "password") { $device.password } else { "" }
                    
                    $requestBodyObj = [PSCustomObject]@{
                        parameters    = $parametersArray
                        sample_period = $samplePeriod
                        credentials   = [PSCustomObject]@{
                            username = $username
                            password = $password
                        }
                    }
                    
                    $credMsg = "[INFO] Including credentials in API request (username: $username)"
                    Write-Host $credMsg -ForegroundColor Gray
                    $credMsg | Out-File -FilePath $logFile -Append
                }
                else {
                    $requestBodyObj = [PSCustomObject]@{
                        parameters    = $parametersArray
                        sample_period = $samplePeriod
                    }
                }
                
                $requestBody = $requestBodyObj | ConvertTo-Json -Depth 100 -Compress:$false
                
                # Debug: Log the actual JSON being sent (only if -debug parameter is set)
                if ($debug) {
                    # Create a masked copy for debug output (mask SECRET_TEXT parameters)
                    $maskedParametersArray = @()
                    foreach ($param in $parametersArray) {
                        # Find if this parameter is SECRET_TEXT
                        $isSecretParam = $false
                        foreach ($key in $parameterMapping.Keys) {
                            $paramInfo = $parameterMapping[$key]
                            if ($paramInfo.id -eq $param.custom_driver_parameter_id -and $paramInfo.value_type -eq "SECRET_TEXT") {
                                $isSecretParam = $true
                                break
                            }
                        }
                        
                        if ($isSecretParam) {
                            # Mask the secret value
                            $maskedParametersArray += [PSCustomObject]@{
                                custom_driver_parameter_id = $param.custom_driver_parameter_id
                                value                      = "********"
                            }
                        }
                        else {
                            # Keep original value
                            $maskedParametersArray += $param
                        }
                    }
                    
                    # Build masked request body for debug display
                    if ($customDriverDetails.requires_credentials -eq $true) {
                        $maskedRequestBodyObj = [PSCustomObject]@{
                            parameters    = $maskedParametersArray
                            sample_period = $samplePeriod
                            credentials   = [PSCustomObject]@{
                                username = $username
                                password = "********"
                            }
                        }
                    }
                    else {
                        $maskedRequestBodyObj = [PSCustomObject]@{
                            parameters    = $maskedParametersArray
                            sample_period = $samplePeriod
                        }
                    }
                    
                    $maskedRequestBody = $maskedRequestBodyObj | ConvertTo-Json -Depth 100 -Compress:$false
                    $debugMsg = "`n[DEBUG] Actual JSON Request Body being sent to API (SECRET_TEXT parameters masked):`n$maskedRequestBody"
                    Write-Host $debugMsg -ForegroundColor Gray
                    $debugMsg | Out-File -FilePath $logFile -Append
                }
                
                $callMessage = "`n[API CALL] Updating association parameters (ID: $associationId)..."
                Write-Host $callMessage -ForegroundColor Yellow
                $callMessage | Out-File -FilePath $logFile -Append
                
                $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Put -Body $requestBody
                
                $successMessage = "[SUCCESS] Association parameters updated successfully"
                Write-Host $successMessage -ForegroundColor Green
                $successMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "OK"
                $device._messages_ = "Parameters updated successfully"
                
                $script:successCount++
                $script:successDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address), Device ID: $deviceID - Updated Association ID: $associationId"
            }
            catch {
                $errorMessage = "[ERROR] Failed to update association parameters: $_"
                Write-Host $errorMessage -ForegroundColor Red
                $errorMessage | Out-File -FilePath $logFile -Append
                
                $device."_apply-result_" = "Error"
                $device._messages_ = "Update operation failed: $_"
                
                $script:failureCount++
                $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - Update Error: $_"
            }
        }
        else {
            # Unsupported operation type
            $unsupportedMsg = "[ERROR] Unsupported operation type: '$operationType' (Collector: $($device.collector_id), IP: $($device.ip_address))"
            Write-Host $unsupportedMsg -ForegroundColor Red
            $unsupportedMsg | Out-File -FilePath $logFile -Append
            
            $device."_apply-result_" = "Error"
            $device._messages_ = "Unsupported operation type: $operationType"
            
            $script:failureCount++
            $script:failureDetails += "Collector ID: $($device.collector_id), Device IP: $($device.ip_address) - Unsupported operation: $operationType"
        }
    }
    
    # STEP 9: Update existing Excel file with status and formatting
    $saveMessage = "`n=== Updating Excel File with Status ===`n"
    Write-Host $saveMessage -ForegroundColor Cyan
    $saveMessage | Out-File -FilePath $logFile -Append
    
    try {
        # Check if ImportExcel module is available for advanced formatting
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            
            $debugMsg = "Opening existing Excel file to update _apply-result_ and _messages_ columns..."
            Write-Host $debugMsg
            $debugMsg | Out-File -FilePath $logFile -Append
            
            # Open the existing Excel file
            $excelPackage = Open-ExcelPackage -Path $excelPath
            $worksheet = $excelPackage.Workbook.Worksheets[1]
            
            $debugMsg = "Excel worksheet has $($worksheet.Dimension.Rows) rows and $($worksheet.Dimension.Columns) columns"
            Write-Host $debugMsg
            $debugMsg | Out-File -FilePath $logFile -Append
            
            # Find _apply-result_ and _messages_ column indices
            $statusColIndex = 0
            $messageColIndex = 0
            
            for ($col = 1; $col -le $worksheet.Dimension.Columns; $col++) {
                $headerValue = $worksheet.Cells[1, $col].Value
                if ($headerValue -eq "_apply-result_") {
                    $statusColIndex = $col
                    $debugMsg = "Found _apply-result_ column at index $col"
                    Write-Host $debugMsg
                    $debugMsg | Out-File -FilePath $logFile -Append
                }
                if ($headerValue -eq "_messages_") {
                    $messageColIndex = $col
                    $debugMsg = "Found _messages_ column at index $col"
                    Write-Host $debugMsg
                    $debugMsg | Out-File -FilePath $logFile -Append
                }
            }
            
            if ($statusColIndex -eq 0 -or $messageColIndex -eq 0) {
                throw "_apply-result_ or _messages_ column not found in Excel file"
            }
            
            # Update cells with data from $devices array
            $updatedCount = 0
            for ($i = 0; $i -lt $devices.Count; $i++) {
                $device = $devices[$i]
                $excelRow = $i + 2  # +2 because Excel is 1-based and row 1 is header
                
                # Update _apply-result_ cell
                if ($device.PSObject.Properties.Name -contains "_apply-result_") {
                    $statusCell = $worksheet.Cells[$excelRow, $statusColIndex]
                    $statusValue = $device."_apply-result_"
                    
                    if (-not [string]::IsNullOrWhiteSpace($statusValue)) {
                        $statusCell.Value = $statusValue
                        
                        if ($statusValue -eq "OK") {
                            # Green and Bold for OK
                            $statusCell.Style.Font.Bold = $true
                            $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                            $updatedCount++
                        }
                        elseif ($statusValue -eq "Error") {
                            # Red and Bold for Error
                            $statusCell.Style.Font.Bold = $true
                            $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                            $updatedCount++
                        }
                        elseif ($statusValue -eq "Skipped") {
                            # Orange and Bold for Skipped
                            $statusCell.Style.Font.Bold = $true
                            $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::Orange)
                            $updatedCount++
                        }
                        elseif ($statusValue -eq "Script already applied") {
                            # Green and Bold for Script already applied
                            $statusCell.Style.Font.Bold = $true
                            $statusCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                            $updatedCount++
                        }
                    }
                }
                
                # Update _messages_ cell
                if ($device.PSObject.Properties.Name -contains "_messages_") {
                    $messageCell = $worksheet.Cells[$excelRow, $messageColIndex]
                    $messageValue = $device._messages_
                    
                    if (-not [string]::IsNullOrWhiteSpace($messageValue)) {
                        $messageCell.Value = $messageValue
                        
                        # Color message based on status
                        if ($device."_apply-result_" -eq "OK") {
                            # Green for success
                            $messageCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                        }
                        elseif ($device."_apply-result_" -eq "Error") {
                            # Red for errors
                            $messageCell.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                        }
                        elseif ($device."_apply-result_" -eq "Skipped") {
                            # Orange for skipped (not bold)
                            $messageCell.Style.Font.Color.SetColor([System.Drawing.Color]::Orange)
                        }
                        elseif ($device."_apply-result_" -eq "Script already applied") {
                            # Green for script already applied
                            $messageCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))  # Dark Green
                        }
                    }
                }
            }
            
            $debugMsg = "Updated $updatedCount status cells in Excel"
            Write-Host $debugMsg -ForegroundColor Green
            $debugMsg | Out-File -FilePath $logFile -Append
            
            # Save and close the Excel file
            $excelPackage.Save()
            Close-ExcelPackage $excelPackage -NoSave
            
            $successMsg = "[OK] Excel file updated successfully with formatting at: $excelPath"
            Write-Host $successMsg -ForegroundColor Green
            $successMsg | Out-File -FilePath $logFile -Append
        }
        else {
            # Fallback: Save as CSV without formatting
            $csvPath = $excelPath -replace '\.xlsx$', '_updated.csv'
            $devices | Export-Csv -Path $csvPath -NoTypeInformation -Force
            $warningMsg = "[WARNING] ImportExcel module not available. Saved updated data as CSV at: $csvPath"
            Write-Host $warningMsg -ForegroundColor Yellow
            $warningMsg | Out-File -FilePath $logFile -Append
        }
    }
    catch {
        $errorMsg = "[ERROR] Failed to update Excel file: $_ | $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFile -Append
        
        # Try to display more debug info
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        
        # Open Excel file even on error to help user review
        if ([string]::IsNullOrEmpty($fileName)) {
            $fileName = $DEFAULT_EXCEL_FILENAME
        }
        if (-not $fileName.EndsWith(".xlsx")) {
            $fileName = "$fileName.xlsx"
        }
        $excelPath = Join-Path $PSScriptRoot $fileName
        
        if (Test-Path $excelPath) {
            $openMsg = "`nOpening Excel file for review..."
            Write-Host $openMsg -ForegroundColor Yellow
            $openMsg | Out-File -FilePath $logFile -Append
            Start-Process $excelPath
        }
    }
    
    # TROUBLESHOOTING: Display Parameter Mapping (only if debug mode)
    if ($debug) {
        $troubleshootingHeader = @"

================================================================================
                    PARAMETER MAPPING (TROUBLESHOOTING)                        
================================================================================
"@
        Write-Host $troubleshootingHeader -ForegroundColor Magenta
        $troubleshootingHeader | Out-File -FilePath $logFile -Append
        
        $mappingOutput = @"

Excel Column Headers (All):
$($columnHeaders -join ', ')

Script Parameters (Excluding collector_id, ip_address, and _columns_):
$($script_parameters_name -join ', ')

Parameter Mapping (Excel Column -> Custom Driver Parameter):
"@
        Write-Host $mappingOutput -ForegroundColor Cyan
        $mappingOutput | Out-File -FilePath $logFile -Append
        
        if ($parameterMapping.Count -gt 0) {
            foreach ($key in $parameterMapping.Keys | Sort-Object) {
                $param = $parameterMapping[$key]
                # Show parsed parameter name if different from Excel column
                if ($param.name -and $param.name -ne $key) {
                    $mappingLine = "  [OK] Excel Column: '$key' -> Parameter: '$($param.name)' | ID: $($param.id) | Type: $($param.value_type) | Description: $($param.description)"
                }
                else {
                    $mappingLine = "  [OK] '$key' -> Parameter ID: $($param.id) | Type: $($param.value_type) | Description: $($param.description)"
                }
                Write-Host $mappingLine -ForegroundColor Green
                $mappingLine | Out-File -FilePath $logFile -Append
            }
        }
        else {
            $noMappingMsg = "  No parameter mappings found"
            Write-Host $noMappingMsg -ForegroundColor Yellow
            $noMappingMsg | Out-File -FilePath $logFile -Append
        }
        
        # Show unmapped Excel columns
        $unmappedExcelColumns = @()
        foreach ($excelParam in $script_parameters_name) {
            if (-not $parameterMapping.ContainsKey($excelParam)) {
                $unmappedExcelColumns += $excelParam
            }
        }
        
        if ($unmappedExcelColumns.Count -gt 0) {
            $unmappedMsg = "`nUnmapped Excel Columns (no corresponding driver parameter):"
            Write-Host $unmappedMsg -ForegroundColor Yellow
            $unmappedMsg | Out-File -FilePath $logFile -Append
            foreach ($col in $unmappedExcelColumns) {
                $colMsg = "  [WARNING] $col"
                Write-Host $colMsg -ForegroundColor Yellow
                $colMsg | Out-File -FilePath $logFile -Append
            }
        }
        
        # Show unmapped driver parameters
        if ($customDriverDetails.parameters) {
            $unmappedDriverParams = @()
            foreach ($driverParam in $customDriverDetails.parameters) {
                $isMapped = $false
                foreach ($key in $parameterMapping.Keys) {
                    if ($parameterMapping[$key].id -eq $driverParam.id) {
                        $isMapped = $true
                        break
                    }
                }
                if (-not $isMapped) {
                    $unmappedDriverParams += $driverParam
                }
            }
            
            if ($unmappedDriverParams.Count -gt 0) {
                $unmappedDriverMsg = "`nUnmapped Driver Parameters (no corresponding Excel column):"
                Write-Host $unmappedDriverMsg -ForegroundColor Yellow
                $unmappedDriverMsg | Out-File -FilePath $logFile -Append
                foreach ($param in $unmappedDriverParams) {
                    $paramMsg = "  [WARNING] $($param.name) (ID: $($param.id), Type: $($param.value_type))"
                    Write-Host $paramMsg -ForegroundColor Yellow
                    $paramMsg | Out-File -FilePath $logFile -Append
                }
            }
        }
    }
    
    # Summary - Build dynamically based on results with colors
    $summaryHeader = @"

================================================================================
                           OPERATION SUMMARY                                    
================================================================================

Operation: $operation
Script Name: $scriptName
Script ID: $scriptID
Total Devices Processed: $script:totalAttempts
Successful: $script:successCount
Failed: $script:failureCount
Skipped: $script:skippedCount

"@
    
    # Write header (no color)
    Write-Host $summaryHeader
    $summaryHeader | Out-File -FilePath $logFile -Append
    
    # Write Successful Details in GREEN
    Write-Host "Successful Details:" -ForegroundColor Green
    "Successful Details:" | Out-File -FilePath $logFile -Append
    if ($script:successDetails.Count -gt 0) {
        $script:successDetails | ForEach-Object { 
            Write-Host "  - $_" -ForegroundColor Green
            "  - $_" | Out-File -FilePath $logFile -Append
        }
    }
    else {
        Write-Host "  (none)" -ForegroundColor Green
        "  (none)" | Out-File -FilePath $logFile -Append
    }
    
    # Only show failed details if there are failures - in RED
    if ($script:failureCount -gt 0) {
        Write-Host ""
        "" | Out-File -FilePath $logFile -Append
        Write-Host "Failed Details:" -ForegroundColor Red
        "Failed Details:" | Out-File -FilePath $logFile -Append
        $script:failureDetails | ForEach-Object { 
            Write-Host "  - $_" -ForegroundColor Red
            "  - $_" | Out-File -FilePath $logFile -Append
        }
    }
    
    # Only show skipped details if there are skipped rows - in YELLOW
    if ($script:skippedCount -gt 0) {
        Write-Host ""
        "" | Out-File -FilePath $logFile -Append
        Write-Host "Skipped Details (missing parameters - may be intentional):" -ForegroundColor Yellow
        "Skipped Details (missing parameters - may be intentional):" | Out-File -FilePath $logFile -Append
        $script:skippedDetails | ForEach-Object { 
            Write-Host "  - $_" -ForegroundColor Yellow
            "  - $_" | Out-File -FilePath $logFile -Append
        }
    }
    
    $summaryFooter = @"

================================================================================
"@
    Write-Host $summaryFooter
    $summaryFooter | Out-File -FilePath $logFile -Append
    
    # Auto-open the Excel file after bulk-apply completes
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = $DEFAULT_EXCEL_FILENAME
    }
    if (-not $fileName.EndsWith(".xlsx")) {
        $fileName = "$fileName.xlsx"
    }
    $excelPath = Join-Path $PSScriptRoot $fileName
    
    if (Test-Path $excelPath) {
        $openMsg = "`nOpening Excel file to review results..."
        Write-Host $openMsg -ForegroundColor Cyan
        $openMsg | Out-File -FilePath $logFile -Append
        
        Start-Process $excelPath
    }
}

# Main execution logic based on operation
switch ($operation) {
    "list-scripts-parameters" {
        # No additional parameters required
        List-Scripts-Parameters
    }
    "open-excel" {
        # No mandatory parameters
        Open-Excel -fileName $filename
    }
    "create-excel" {
        # Validate required parameters
        if ([string]::IsNullOrEmpty($script_name)) {
            Write-Host "ERROR: -script_name parameter is mandatory for create-excel operation!" -ForegroundColor Red
            Show-Usage
        }
        if ([string]::IsNullOrEmpty($collector_ids)) {
            Write-Host "ERROR: -collector_ids parameter is mandatory for create-excel operation!" -ForegroundColor Red
            Show-Usage
        }
        Create-Excel -scriptName $script_name -collectorIds $collector_ids -fileName $filename
    }
    "bulk-apply" {
        # Validate required parameters
        if ([string]::IsNullOrEmpty($script_name)) {
            Write-Host "ERROR: -script_name parameter is mandatory for bulk-apply operation!" -ForegroundColor Red
            Show-Usage
        }
        bulk-Apply-Script -scriptName $script_name -fileName $filename
    }
    default {
        Write-Host "ERROR: Invalid operation specified!" -ForegroundColor Red
        Show-Usage
    }
}

$logMessage = "`nLOG FILE: $logFile"
Write-Host $logMessage
$logMessage | Out-File -FilePath $logFile -Append

