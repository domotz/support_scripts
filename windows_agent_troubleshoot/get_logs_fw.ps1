# Domotz Agent Windows troubleshoot script
# What it does: 
# - gets Windows OS info
# - get Domotz Agent Logs
# - checks for Domotz Cloud connectivity (outgoing)
# - perform a test with the selected Speedtest - if enabled -

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
Write-Host "This is the Domotz Support Diagnostic application. 
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

# Check for Domotz service
if (-not(Get-Service $domotzService -ErrorAction SilentlyContinue))
{
    Write-Host ""
    Write-Host "CRITICAL: $domotzService does not exist! Please, send this output to support@domotz.com"
    Write-Host ""
    exit
}

Write-Host "Loading..."
Write-Host ""

# Test installation dir
if (-not(Test-Path -Path $daemonLogDir)) {
    try {
        Write-Host ""
        Write-Host "WARNING: $daemonLogDir is missing. The Domotz agent has not been installed in this path. Please adjust the deamonLogDir and domotzNode paths variables at the top of this script, or contact support@domotz.com"
        Write-Host ""
        exit
    }
    catch {
        throw $_.Exception.Message
    }
}

# Getting Agent Id
if (-not(Test-Path -Path $agentConfFile -PathType Leaf)) {
    try {
        Write-Host ""
        Write-Host "WARNING: $agentConfFile is missing. This agent has not been registered to an account - please contact support@domotz.com"
        Write-Host ""
        exit
    }
    catch {
        throw $_.Exception.Message
    }
}

$agentObj=Get-Content -Raw -Path $agentDataDir\domotz.json | ConvertFrom-Json
$agentID= $agentObj | Select-Object -ExpandProperty "id"
if (!$agentID) {
    Write-Host ""
    Write-Host "WARNING: The Agent ID is missing, pls check $agentConfFile contents. This agent has not been registered to an account - please contact support@domotz.com"
    Write-Host ""
    exit
}
else {
    $customerLogDir="agent-$agentID-Logs-$date"
    $reportFile="$DesktopPath\$customerLogDir\agent_short_report.txt"

    # Creating Log dir
    if (!(Test-Path $DesktopPath\$customerLogDir -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir | Out-Null
    }
    Add-Content $reportFile "--Agent Short Report $date--"
    Add-Content $reportFile ""
    Add-Content $reportFile "[Agent Details]"
}

# Getting Agent Name
$agentName=$agentObj | Select-Object -ExpandProperty "display_name"
if (!$agentName) {
    Write-Host ""
    Write-Host "WARNING: The Agent Name is missing, pls check $agentConfFile contents. - please contact support@domotz.com"
    Write-Host ""
    exit
}
else {
    $customerLogDir="agent-$agentID-Logs-$date"
    $reportFile="$DesktopPath\$customerLogDir\agent_short_report.txt"
}

# Get Domotz Agent service properties 
$domotzServiceFile="$DesktopPath\$customerLogDir\DomotzServiceProperties.txt"
Get-Service $domotzService | select Displayname,Status | Out-File $domotzServiceFile

# Get Agent versions
$agentVersion=$agentObj | Select-Object -ExpandProperty "version" | Select-Object -ExpandProperty "agent"
$agentPkgVersion=$agentObj | Select-Object -ExpandProperty "version" | Select-Object -ExpandProperty "package"

# Get MAC Address
$agentMACAddr=$agentObj | Select-Object -ExpandProperty "licence" | Select-Object -ExpandProperty "bound_mac_address"

# Get address the agents listens on in the network
$agentListensOn=$agentObj | Select-Object -ExpandProperty "listen_on" 

# Get Computer info
$osinfoFile="$DesktopPath\$customerLogDir\os_info.txt"
Get-ComputerInfo | Out-File $osinfoFile 

# Cell Definition - wrinting Cell in report and set the hosts to be checked by region
$messaging_host=$agentObj | Select-Object -ExpandProperty "message_broker" | Select-Object -ExpandProperty "host"

if (!$messaging_host) {
    Write-Host ""
    Write-Host "WARNING: The Agent CELL is missing, pls check $agentConfFile contents. - please contact support@domotz.com"
    Write-Host ""
    exit
}

else {
    if ($messaging_host -like '*us*') {
        $cell="US"
        $hosts=$ushosts
    }
    if ($messaging_host -like '*eu*') {
        $cell="EU"
        $hosts=$euhosts
    }
    Add-Content $reportFile "Agent cell: $cell"
    Add-Content $reportFile "Agent ID: $agentID"
    Add-Content $reportFile "Agent Name: $agentName"
    Add-Content $reportFile "Agent version: $agentVersion"
    Add-Content $reportFile "Agent pkg version: $agentPkgVersion"
    Add-Content $reportFile "Agent MAC: $agentMACAddr"
    Add-Content $reportFile "Agent listens on: $agentListensOn"
}

# Collect Network Information
$netInfo=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/net-info
if (!$netInfo) {
    Add-Content $reportFile "WARNING: Unable to get net-info logs from 127.0.0.1:3000"
}
else {
    $netinfoFile="$DesktopPath\$customerLogDir\net_info.txt"
    Add-Content $netinfoFile "--Agent NetFinfo Report $date--"
    $netInfo.Content | ConvertFrom-Json | ConvertTo-Json -depth 100 | Out-File $netinfoFile
}

# Create short report
$loggingVersion=$agentObj | Select-Object -ExpandProperty "conf_version"
if (!$loggingVersion) {
    $loggingVersionType="old_logging"
    Add-Content $reportFile "LoggingType=$loggingVersionType"

}
else {
    $loggingVersionType="new"
    Add-Content $reportFile "LoggingType=$loggingVersionType"
    Add-Content $reportFile "LoggingVersion=$loggingVersion"
    
}

# Collect Listener logs
Write-Host ""
Write-Host -noNewLine "-> Collecting Domotz Logs... please wait..."
if (!(Test-Path $DesktopPath\$customerLogDir\listener_logs -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir\listener_logs | Out-Null
}
Copy-Item -Path $listernerLogDir\* -Destination $DesktopPath\$customerLogDir\listener_logs | Out-Null

# Collect Daemon logs
if (!(Test-Path $DesktopPath\$customerLogDir\daemon_logs -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir\daemon_logs | Out-Null
}
Copy-Item -Path $daemonLogDir\domotzagent.err.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null
Copy-Item -Path $daemonLogDir\domotzagent.out.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null
Copy-Item -Path $daemonLogDir\domotzagent.wrapper.log -Destination $DesktopPath\$customerLogDir\daemon_logs | Out-Null

# Collect flush logs
$flushLog=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/log/flush
if (!$flushLog) {
    Add-Content $reportFile "WARNING: Unable to flush logs from 127.0.0.1:3000"
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
    Add-Content $npcapIssueReport "### NB: This has to be reviewed and could not be accurate! have to find better evidence $match_str is too generic!!"
    Select-String -Path $DesktopPath\$customerLogDir\listener_logs\*.log.* -Patter $match_str | Out-File -Encoding Ascii -Append $npcapIssueReport
}
Write-Host " Done!"


##ADD check for the Nmap version and stuff - not ready yet ...
##https://domotzjira.atlassian.net/browse/NI-386 
$domotzStatus=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/status
$domotzStatusObj=$domotzStatus | ConvertFrom-Json

# $nmapVersion=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "version"
# $npcapVersion=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "npcap"
# $nmapLiblua=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "nmap-liblua"
# $nmapLibssh2=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "nmap-libssh2"
# $libPcap=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "libpcap"
# $ipv6Support=$domotzStatusObj | Select-Object -ExpandProperty "package" | Select-Object -ExpandProperty "nmap" | Select-Object -ExpandProperty "libraries" | Select-Object -ExpandProperty "ipv6"

# Add-Content $reportFile "[Nmap Details]"
# Add-Content $reportFile "Nmap version=$nmapVersion"
# Add-Content $reportFile "Npcap version=$npcapVersion"
# Add-Content $reportFile "NmapLibLua version=$nmapLiblua"
# Add-Content $reportFile "NpmapLibSsh2 version=$nmapLibssh2"
# Add-Content $reportFile "lib-pcap version=$libPcap"
# Add-Content $reportFile "nmap ipv6Support=$ipv6Support"

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
    $speedtestReportFile="$DesktopPath\$customerLogDir\fast_test.txt"
    &"$domotzNode" $currentDir\fast_speed_test.js | Out-File $speedtestReportFile
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
Read-Host -Prompt "Domotz Diagnostics has finished his job! Thank you for using it!
Please, press any key to EXIT!" 