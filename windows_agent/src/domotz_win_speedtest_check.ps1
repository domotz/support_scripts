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
$dscriptver="1.1speedtest"
# Changelog
# v1.1 added various sections and when passed to node 20 - renamed domotz-remote-pawn dir into domotz-remote-pawn-ng 


# Check if you have administrative privileges to run this script
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this utility. You should run this with administrative privileges."
    Read-Host -Prompt "please press ENTER or CTRL+C to quit"
    break
}
else {
    Write-Information "Code is running as administrator - nice to hear that!"
}

# Check if requirements are met (OS version, Ps version)
$os = Get-CimInstance Win32_OperatingSystem
if ($os.Version -ge 10.0) {
    Write-Host "OS version check --> Passed..."
}
else {
    $osVerCaption = (Get-WMIObject win32_operatingsystem) | Select-Object -Expandproperty Caption
    Write-Warning "Your OS version: ----->  $osVerCaption $osVer <------ is not supported by this script. SSH server should be installed manually."
    break
}



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
|    Monitor any network and IT infrastructure   |
+------------------------------------------------+
WARNING: SPEEDTEST CHECK ONLY!!!
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
