# Domotz Agent Windows troubleshoot script
# What it does: 
# - gets Windows OS info
# - gets node interfaces and gateway info
# - get Domotz Agent Logs
# - checks for network interfaces with more than one ip address
# - check for npcap issue
# - checks for Domotz Cloud connectivity (outgoing)
# - perform a test with the selected Speedtest - if enabled -
# TODO -- add progression messages for each check/task made.
$dscriptver="1.0"

# Motd
Write-Host "
+------------------------------------------------+
|  ___                             _             |
| (  _'\                          ( )_           |
| | | ) |   _     ___ ___     _   | ,_) ____     |
| | | | ) /'_'\ /' _ ' _ '\ /'_'\ | |  (_  ,)    |
| | |_) |( (_) )| ( ) ( ) |( (_) )| |_  /'/_     |
| (____/''\___/'(_) (_) (_)'\___/''\__)(____)    |
| ---------------------------------------------- |
| The RMM tool for Networks and Connected Devices|
+------------------------------------------------+
"
Write-Host "This is the Domotz Diagnostic application. 
It will create a zip file on your Desktop which you will send to support@domotz.com
"

Read-Host -Prompt "Press ENTER to continue or CTRL+C to quit" 

$agentInstDir_compl= Get-ItemProperty HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\domotz | Select-Object UninstallString
$agentInstDir=$agentInstDir_compl.UninstallString.Trim('"') -replace "uninstall.exe", ""
$agentDataDir="$Env:ALLUSERSPROFILE\domotz"

# Domotz logs variables
$date=Get-Date -Format "dd-MM-yyyy-HH-mm-ss"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$agentConfFile="$agentDataDir\domotz.json"
$listernerLogDir="$agentDataDir\log"
$daemonLogDir="$agentInstDir\bin\daemon"
$domotzService = "Domotz Agent"
$domotzNode="$agentInstDir\bin\domotz_node.exe"
$currentDir=$PSScriptRoot
$customerLogDir="domotz-diagnostics-$date"

# Domotz Firewall hosts and ports variables
$ProgressPreference="SilentlyContinue";
# Get DNS settings
$googleDns="8.8.8.8"
$dnsServers = Get-DnsClientServerAddress | findstr -i $googleDns
$echoHost = "echo.domotz.com"
# Domotz Cloud Hosts and Ports
$ushosts= @(
    [pscustomobject]@{host = "portal.domotz.com"; port = "443"; region = "EU and US and APAC"}
    [pscustomobject]@{host = "api-us-east-1-cell-1.domotz.com"; port = "443"; region = "US and APAC"}
    [pscustomobject]@{host = "messaging-us-east-1-cell-1.domotz.com"; port = "5671"; region = "US"}
)
$euhosts = @(
    [pscustomobject]@{host = "portal.domotz.com"; port = "443"; region = "EU and US and APAC"}
    [pscustomobject]@{host = "api-eu-west-1-cell-1.domotz.com"; port = "443"; region = "EU and APAC"}
    [pscustomobject]@{host = "messaging-eu-west-1-cell-1.domotz.com"; port = "5671"; region = "EU and APAC"}
)
# Domotz Boxes 
$domotzBox_hosts = @(
    [pscustomobject]@{host = "provisioning.domotz.com"; port = "4505"; model = "B-01,B-03,B-11"}
    [pscustomobject]@{host = "provisioning.domotz.com"; port = "4506"; model = "B-01,B-03,B-11"}
    [pscustomobject]@{host = "messaging.orchestration.domotz.com"; port = "5671"; model = "B-12"}
    [pscustomobject]@{host = "api.orchestration.domotz.com"; port = "443"; model = "B-12"}
    [pscustomobject]@{host = "api.snapcraft.io"; port = "443"; model = "B-12"}
)

# Creating Log dir
Write-Host ""
Write-Host -noNewLine "-> Creating diagnostic/logs dirs on Desktop"

if (!(Test-Path $DesktopPath\$customerLogDir -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir | Out-Null
}

# Creating short report file
$reportFile="$DesktopPath\$customerLogDir\agent_short_report.txt"
Add-Content $reportFile "--Agent Short Report $date--"
Add-Content $reportFile ""
Add-Content $reportFile "[Agent Details]"

# Difining warnings report file
$warningsFile="$DesktopPath\$customerLogDir\warnings_report.txt"

# Test if the Domotz agent is installed (in the standards paths)
if (-not(Test-Path -Path $daemonLogDir)) {
    try {
        $msg_warn_dlog_missing="WARNING: $daemonLogDir is missing. The Domotz agent has not been installed in this path. If the agent is running, you might want to adjust the deamonLogDir and domotzNode paths variables at the top of this script."
        Write-Host ""
        Write-Host $msg_warn_dlog_missing
        Write-Host ""
        Add-Content $warningsFile ""
        Add-Content $warningsFile $msg_warn_dlog_missing
    }
    catch {
        throw $_.Exception.Message
    }
}
Write-Host " Done!"


# Check Domotz Agent installation/service status
Write-Host ""
Write-Host -noNewLine "-> Checking Domotz Agent installation/service status"

# Check if the Domotz Service exists
if (-not(Get-Service $domotzService -ErrorAction SilentlyContinue))
{
    $msg_warn_dservice_missing="CRITICAL: $domotzService does not exist!"
    Write-Host ""
    Write-Host $msg_warn_dservice_missing
    Write-Host ""
    Add-Content $warningsFile ""
    Add-Content $warningsFile $msg_warn_dservice_missing
    exit
}
else {
    # Get Domotz Agent service status 
    $domotzServiceStatus=Get-Service $domotzService | Select-Object -ExpandProperty Status

    if ($domotzServiceStatus -ne "Running") {
        Add-Content $warningsFile ""
        Add-Content $warningsFile "Domotz Agent Service is not running or not running properly"
        Add-Content $warningsFile "-> Domotz Service status:"
        Add-Content $warningsFile $domotzServiceStatus
    }
    else {
        Add-Content $reportFile ""
        Add-Content $reportFile "-> Domotz Service status:"
        Add-Content $reportFile $domotzServiceStatus
    }
}

# Getting Agent Id
if (-not(Test-Path -Path $agentConfFile -PathType Leaf)) {
        Add-Content $warningsFile ""
        Add-Content $warningsFile "WARNING: $agentConfFile is missing. This agent has not been registered to an account."
    }
else {
    $agentObj=Get-Content -Raw -Path $agentDataDir\domotz.json | ConvertFrom-Json
    
    if (!$agentObj) {
        Add-Content $warningsFile ""
        Add-Content $warningsFile "WARNING: There is something from in the $agentConfFile File. Please check its contents."
        exit
    }
    else {
        # Getting Agent parameters from domotz.json file
        
        ## Calculating Agent Cell (US or EU/ROW)
        $messaging_host=$agentObj | Select-Object -ExpandProperty "message_broker" | Select-Object -ExpandProperty "host"
        if ($messaging_host -like '*us*') {
            $cell="US"
            $hosts=$ushosts
        }
        if ($messaging_host -like '*eu*') {
            $cell="EU"
            $hosts=$euhosts
        }
        $agentID= $agentObj | Select-Object -ExpandProperty "id"
        $agentName=$agentObj | Select-Object -ExpandProperty "display_name"
        # Get Agent versions
        $agentVersion=$agentObj | Select-Object -ExpandProperty "version" | Select-Object -ExpandProperty "agent"
        $agentPkgVersion=$agentObj | Select-Object -ExpandProperty "version" | Select-Object -ExpandProperty "package"
        $agentMACAddr=$agentObj | Select-Object -ExpandProperty "licence" | Select-Object -ExpandProperty "bound_mac_address"
        # Get address the agent is listening on in the network
        $agentListensOn=$agentObj | Select-Object -ExpandProperty "listen_on" 

        Add-Content $reportFile "Diagnostic Script version: $dscriptver"
        Add-Content $reportFile "Agent cell: $cell"
        Add-Content $reportFile "Agent ID: $agentID"
        Add-Content $reportFile "Agent Name: $agentName"
        Add-Content $reportFile "Agent version: $agentVersion"
        Add-Content $reportFile "Agent pkg version: $agentPkgVersion"
        Add-Content $reportFile "Agent MAC: $agentMACAddr"
        Add-Content $reportFile "Agent listens on: $agentListensOn"
    }
}
Write-Host " Done!"

# Collect Network Information
Write-Host ""
Write-Host -noNewLine "-> Collecting Network Configuration info...."

try {
    $netInfo=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/net-info -TimeoutSec 10
}
catch {
    Write-Host "Collecting Network Information -->" $_.Exception.Message
    $lastError=$_.Exception.Message
}

if (!$netInfo) {
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> WARNING: Unable to get net-info logs from 127.0.0.1:3000 " 
    Add-Content $warningsFile $lastError
}
else {
    $netinfoFile="$DesktopPath\$customerLogDir\net_info.txt"
    Add-Content $netinfoFile "--Agent NetFinfo Report $date--"
    Add-Content $netinfoFile ""
    $netInfo.Content | ConvertFrom-Json | ConvertTo-Json -depth 100 | Out-File $netinfoFile
}

# Get Interfaces from Domotz Node
if (Test-Path -Path $domotzNode){
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir\node_info | Out-Null
    $ipDomotzNodeInt="$DesktopPath\$customerLogDir\node_info\node_interfaces.txt"
    &"$domotzNode" -e "console.log(require(`'os`').networkInterfaces());" | Out-File $ipDomotzNodeInt
}
else {
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> WARNING:  $domotzNode is not available" 
}

# Get Node deafultgw
if (Test-Path -Path $domotzNode){
    $ipDomotzNodeGw="$DesktopPath\$customerLogDir\node_info\node_gw_info.txt"
    $ipDomotzNodeGw2="$DesktopPath\$customerLogDir\node_info\node_gw_info_2.txt"
    $ipDomotzNodeGw3="$DesktopPath\$customerLogDir\node_info\node_gw_info_3.txt"
    Set-Location $remotePawnDir
    &"$domotzNode" -e "console.log(JSON.stringify(require(`'default-gateway`').v4.sync()))" | Out-File $ipDomotzNodeGw
    Get-CimInstance Win32_NetworkAdapterConfiguration -filter "IPEnabled=true" | Select-Object DefaultIPGateway,Index | ConvertTo-JSON | Out-File $ipDomotzNodeGw2
    $gwIndex = Get-CimInstance Win32_NetworkAdapterConfiguration -filter "IPEnabled=true" |  Select-Object DefaultIPGateway,Index |Where-Object { $_.DefaultIPGateway -ne $null} | Select-Object -ExpandProperty Index
    Get-CimInstance Win32_NetworkAdapter -filter Index=$gwIndex | Select-Object NetConnectionID,MacAddress | ConvertTo-JSON | Out-File $ipDomotzNodeGw3
}

# Get Interfaces IP Info
$ipNicInfoFile="$DesktopPath\$customerLogDir\interfaces_brief_info.txt"
Get-NetIPConfiguration -All | Out-File $ipNicInfoFile

# Check for interface with double ips
$getInt= Get-NetIPAddress -AddressFamily IPv4 | Select-Object -ExpandProperty ifIndex 
$getInt_unique=$getInt | Select-Object -Unique
$dup_int=Compare-Object -ReferenceObject $getInt_unique -DifferenceObject $getInt | Select-Object -ExpandProperty InputObject

if ($dup_int) {
    $dup_intName=Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.ifIndex -EQ $dup_int } | Select-Object -ExpandProperty InterfaceAlias |  Select-Object -Unique 
    $dup_IpAddr=Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.ifIndex -EQ $dup_int } | Select-Object -ExpandProperty IPAddress 
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> CRITICAL:  More that one ip attached to the same Network Interface"
    Add-Content $warningsFile $dup_intName
    Add-Content $warningsFile $dup_IpAddr
}

# Get Computer info
Write-Host ""
Write-Host -noNewLine "-> Collecting Operating System info... please wait..."

$osinfoFile="$DesktopPath\$customerLogDir\os_info.txt"
Get-ComputerInfo | Out-File $osinfoFile

Write-Host " Done!"

# Collect Listener logs
Write-Host ""
Write-Host -noNewLine "-> Collecting Domotz Logs... please wait..."

if (Test-Path $listernerLogDir -PathType Container){
    if (!(Test-Path $DesktopPath\$customerLogDir\listener_logs -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir\listener_logs | Out-Null
    }
    Copy-Item -Path $listernerLogDir\* -Destination $DesktopPath\$customerLogDir\listener_logs | Out-Null
}
# Collect Daemon logs
if (Test-Path $daemonLogDir -PathType Container) {
    if (!(Test-Path $DesktopPath\$customerLogDir\daemon_logs -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir\daemon_logs | Out-Null
    }
    Copy-Item -Path $daemonLogDir\domotzagent.err.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null
    Copy-Item -Path $daemonLogDir\domotzagent.out.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null
    Copy-Item -Path $daemonLogDir\domotzagent.wrapper.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null
}

# Collect flush logs
try {
    $flushLog=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/log/flush -TimeoutSec 10
}
catch {
    Write-Host "Collecting Listener logs -->" $_.Exception.Message
    $lastError=$_.Exception.Message
}

if (!$flushLog) {
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> WARNING: Unable to flush logs from 127.0.0.1:3000"
    Add-Content $warningsFile $lastError
}
else {
    $flushLog | ConvertFrom-Json | Out-File $DesktopPath\$customerLogDir\listener_logs\flushed_log.txt
}

Write-Host " Done!"

# check for Npcap issue - This has to be reviewed - have to find better evidence $match_str is too generic #TODO
Write-Host ""
Write-Host -noNewLine "-> Checking for win Npcap issues please wait..."

$npcap_info=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate | Where -Property DisplayName -Match Npcap
$npcap_version=$npcap_info.DisplayVersion
Add-Content $reportFile "Npcap version=$npcap_version"

$match_str="Cannot find MAC address for device with IP"
if(Select-String -Path $DesktopPath\$customerLogDir\listener_logs\*.log.* -Patter $match_str){
    $npcapIssueReport="$DesktopPath\$customerLogDir\npcap_issue_maybe_detected.txt"
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> WARNING: this is agent can have the NPCAP issue"
    Add-Content $warningsFile "### NB: This has to be reviewed and could not be accurate! have to find better evidence $match_str is too generic!!"
    Select-String -Path $DesktopPath\$customerLogDir\listener_logs\*.log.* -Patter $match_str | Out-File -Encoding Ascii -Append $npcapIssueReport
}
Write-Host " Done!"

##ADD check for the Nmap version and stuff - not ready yet ...
##https://domotzjira.atlassian.net/browse/NI-386 
try {
    $domotzStatus=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/status -TimeoutSec 10
}
catch {
    Write-Host "Checking for the Nmap version and stuff -->" $_.Exception.Message
    $lastError=$_.Exception.Message
}

if (!$domotzStatus) {
    Add-Content $warningsFile ""
    Add-Content $warningsFile "-> WARNING: Unable to check for Nmap version and stuff from 127.0.0.1:3000"
    Add-Content $warningsFile $lastError
}
else {
    $domotzStatusObj=$domotzStatus | ConvertFrom-Json
    $nmapVersion=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "version"
    $npcapVersion=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "npcap"
    $nmapLiblua=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "nmap-liblua"
    $nmapLibssh2=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "nmap-libssh2"
    $libPcap=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "libpcap"
    $ipv6Support=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "ipv6"

    Add-Content $reportFile "[Nmap Details]"
    Add-Content $reportFile "Nmap version=$nmapVersion"
    Add-Content $reportFile "Npcap version=$npcapVersion"
    Add-Content $reportFile "NmapLibLua version=$nmapLiblua"
    Add-Content $reportFile "NpmapLibSsh2 version=$nmapLibssh2"
    Add-Content $reportFile "lib-pcap version=$libPcap"
    Add-Content $reportFile "nmap ipv6Support=$ipv6Support"
}
Write-Host " Done!"


# Domotz Agent -- Test Firewall
# Messages
Write-Host ""
Write-Host -noNewLine "-> Testing network connection to Domotz Cloud... please wait..."
$fwReportFile="$DesktopPath\$customerLogDir\firewall_check_report.txt"
$openonfw="$date - [E] - KO -- Connection Error - Please open on your Firewall for OUTGOING connections to:"
Add-Content $fwReportFile "--Agent Firewall Report $date--"
Add-Content $fwReportFile "[Firewall]"
if (Test-Connection -ComputerName $echoHost -Quiet) { 
    Add-Content $fwReportFile "Ping to echo.domotz.com succeded"
    
}
else {
    Add-Content $fwReportFile "$date - [E] - KO -- Ping to echo.domotz.com unsuccessful"
}
Add-Content $fwReportFile ""
if ($dnsServers) {
    Add-Content $fwReportFile "DNS settings OK"
}
else {
    Add-Content $fwReportFile "[W] - WARNING - Please use Google Public DNS for your Domotz agent host machine! (8.8.8.8 and 8.8.4.4)"
}
Add-Content $fwReportFile ""

foreach ($a in $hosts) {
    $null = (Test-NetConnection -ComputerName $a.host -Port $a.port -ErrorAction SilentlyContinue -ErrorVariable ConnectionError).TcpTestSucceeded
    
    $ahost=$a.host
    $aport=$a.port
    $aregion=$a.region

    if ($ConnectionError) {
        Add-Content $fwReportFile "[W] This is required by Region: $aregion"
        Add-Content $fwReportFile "[E] - KO -- $openonfw $ahost on Port $aport"
        Add-Content $fwReportFile ""
    }
    else {
        Add-Content $fwReportFile "Connection to $ahost - $aport OK"
        Add-Content $fwReportFile ""
    }
}

foreach ($a in $domotzBox_hosts) {
    $null = (Test-NetConnection -ComputerName $a.host -Port $a.port -ErrorAction SilentlyContinue -ErrorVariable ConnectionError).TcpTestSucceeded
    
    $ahost=$a.host
    $aport=$a.port
    $aregion=$a.region
    $amodel=$a.model

    if ($ConnectionError) {
        Add-Content $fwReportFile "[W] This applies only to Domotz Box - $amodel"
        Add-Content $fwReportFile "[E] - KO --$openonfw $ahost on Port $aport"
    }
    else {
        Add-Content $fwReportFile "Connection to $ahost - $aport OK"
        Add-Content $fwReportFile ""
    }
}
Write-Host " Done!"


# Speedtest check
if (Test-Path -Path $currentDir\fast_speed_test.js -PathType Leaf) {
    Write-Host ""
    Write-Host "Speedtest check --- this could take some time --- Please wait... (it is running don't worry :) )"
    
    $speedtestReportFile="$DesktopPath\$customerLogDir\speedtest_check_log.txt"
    &"$domotzNode" $currentDir\fast_speed_test.js | Out-File $speedtestReportFile
    
    Write-Host " Done!"
}

# Create the final Zip file
Write-Host ""
Write-Host -noNewLine "-> Creating zip file on Desktop... please wait..."

Compress-Archive -Path $DesktopPath\$customerLogDir $DesktopPath\$customerLogDir.zip

if (Test-Path -Path $DesktopPath\$customerLogDir.zip){
    Remove-Item -Force -Recurse -Path $DesktopPath\$customerLogDir
}
else {
    Write-Host ""
    Write-Host "WARNING: the zip file could not be created!! - please contact support@domotz.com"
    Write-Host ""
    exit
}
Write-Host " Done!"
Write-Host "
PLEASE READ THIS:"
Write-Host "File $DesktopPath\$customerLogDir.zip file which contains your agent logs and reports has been created!" 
Write-Host "N.B. Please send this to support@domotz.com"
Write-Host ""
Write-Host "Domotz Diagnostics has finished his job! Thank you for using it!"
Write-Host ""
Write-Host "Please, press ENTER to EXIT or close this window!" -NoNewLine
$UserInput = $Host.UI.ReadLine()
