# Required to invoke REST APIs
# Install the required module if not already installed
if (-not (Get-Module -ListAvailable -Name 'PSReadline')) {
    Install-Module -Name 'PSReadline' -Force
}

# Domotz API variables
$apiUrl = "https://api-us-east-1-cell-1.domotz.com/public-api/v1/"
$apiKey = "YOUR-API-KEY"  # Replace with your Domotz API Key

# Filters
$siteNames = Read-Host "Enter site names separated by commas (Leave empty for all sites)"
$deviceType = Read-Host "Enter device type (Leave empty to skip)"
$deviceMake = Read-Host "Enter device make (Leave empty to skip)"
$deviceModel = Read-Host "Enter device model (Leave empty to skip)"
$firmwareVersion = Read-Host "Enter firmware version (Leave empty to skip)"
$filterAgentStatus = Read-Host "Filter agents/sites by status? (online/offline/all)"
$filterDeviceStatus = Read-Host "Filter devices by status? (online/offline/all)"
$filterSNMPStatus = Read-Host "Enter SNMP Status (AUTHENTICATED, NOT_AUTHENTICATED, or leave empty to skip)"
$filterAuthStatus = Read-Host "Enter Authentication Status (Leave empty to skip)"

# API helper function to make authenticated requests
function Invoke-DomotzAPI {
    param (
        [string]$endpoint,
        [string]$method = "GET"
    )
    $headers = @{
        "x-api-key" = $apiKey
    }
    $response = Invoke-RestMethod -Uri "$apiUrl$endpoint" -Method $method -Headers $headers
    return $response
}

# Helper function to extract labels from complex objects
function Extract-Label {
    param (
        [object]$property
    )
    if ($property -is [string]) {
        return $property
    } elseif ($property -is [pscustomobject] -and $property.label) {
        return $property.label
    }
    return ""
}

# Step 1: Retrieve all agents/sites with pagination (Modified Section)
Write-Host "Retrieving all agents/sites..."
$agents = @()
$pageSize = 100  # Max allowed by API
$pageNumber = 0
do {
    $response = Invoke-DomotzAPI "agent?page_size=$pageSize&page_number=$pageNumber"
    $agents += $response
    $pageNumber++
} while ($response.Count -eq $pageSize)  # Continue until we get less than 100 results

# Filter agents based on the site names and status (if provided)
$filteredAgents = $agents | Where-Object {
    ($siteNames -eq "" -or ($siteNames -split ",").Trim() -contains $_.display_name) -and
    ($filterAgentStatus -eq "all" -or $_.status.value -eq $filterAgentStatus)
}

# Step 2: Get devices for each agent (Unchanged)
$devices = @()
foreach ($agent in $filteredAgents) {
    Write-Host "Retrieving devices for site: $($agent.display_name)..."
    $agentDevices = Invoke-DomotzAPI "agent/$($agent.id)/device"
    
    # Add agent and organization information to each device
    foreach ($device in $agentDevices) {
        $organizationName = "No Organization"
        if ($agent.organization.name) {
            $organizationName = $agent.organization.name
        }

        $device | Add-Member -MemberType NoteProperty -Name "Organization" -Value $organizationName
        $device | Add-Member -MemberType NoteProperty -Name "agent_name" -Value $agent.display_name
        $device | Add-Member -MemberType NoteProperty -Name "agent_id" -Value $agent.id
        $device | Add-Member -MemberType NoteProperty -Name "agent_status" -Value $agent.status.value
    }
    $devices += $agentDevices
}

# Step 3: Apply filters to devices (Unchanged)
$filteredDevices = $devices | Where-Object {
    ($deviceType -eq "" -or ($_.type.label -eq $deviceType)) -and
    ($deviceMake -eq "" -or ($_.user_data.vendor -eq $deviceMake)) -and
    ($deviceModel -eq "" -or ($_.user_data.model -eq $deviceModel)) -and
    ($firmwareVersion -eq "" -or ($_.details.firmware_version -eq $firmwareVersion)) -and
    ($filterSNMPStatus -eq "" -or ($_.snmp_status -eq $filterSNMPStatus)) -and
    ($filterAuthStatus -eq "" -or ($_.authentication_status -eq $filterAuthStatus)) -and
    ($filterDeviceStatus -eq "all" -or $_.status -eq $filterDeviceStatus)
}

# Step 4: Output data to CSV (Unchanged)
Write-Host "Generating CSV file..."
$csvData = $filteredDevices | ForEach-Object {
    [PSCustomObject]@{
        'Organization Name' = $_.Organization
        'Site/Agent Name' = $_.agent_name
        'Domotz Site/Agent ID' = $_.agent_id
        'Device Name' = $_.user_data.name
        'Domotz Device ID' = $_.id
        'Device IP Address(es)' = if ($_.ip_addresses) { $_.ip_addresses[0] } else { "Not Available" }
        'Device MAC Address' = if ($_.hw_address) { $_.hw_address } else { "Not Available" }
        'Serial Number' = $_.details.serial
        'Firmware Version' = $_.details.firmware_version
        'Device Type' = Extract-Label $_.type
        'Device Make' = $_.user_data.vendor
        'Device Model' = $_.user_data.model
        'SNMP Status' = $_.snmp_status
        'Authentication Status' = $_.authentication_status
        'Room' = $_.details.room
        'Zone' = $_.details.zone
        'First Seen On' = $_.first_seen_on
        'SNMP Read Community' = $_.details.snmp_read_community
        'SNMP Write Community' = $_.details.snmp_write_community
        'Tags' = if ($_.importance -eq 'VITAL') { 'Important' } else { '' }
        'Protocol' = $_.protocol
        'Device Status' = $_.status
    }
}

$csvFileName = "Domotz_Device_Report.csv"
$csvData | Export-Csv -Path $csvFileName -NoTypeInformation

# Keep the window open for the user to see the results
Write-Host "CSV file generated: $csvFileName"
Write-Host "Press any key to exit..."
[void][System.Console]::ReadKey($true)
exit
