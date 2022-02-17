# Domotz Agent Windows troubleshoot script if not activation occurs (domotz.json not present - 127.0.0.1:3000 does not work)
# What it does: 
# - gets Windows OS info
# - get Domotz Agent Logs

$dscriptver="0.1.1"

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
It will create a zip file on your Desktop which you will send to support@domotz.com - ver. $dscriptver
"

Read-Host -Prompt "Press ENTER to continue or CTRL+C to quit" 


$agentInstDir_compl= Get-ItemProperty HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\domotz | Select-Object UninstallString
$agentInstDir=$agentInstDir_compl.UninstallString.Trim('"') -replace "uninstall.exe", ""
$agentDataDir="$Env:ALLUSERSPROFILE\domotz"

# Domotz logs variables
$date=Get-Date -Format "dd-MM-yyyy-HH-mm-ss"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$listernerLogDir="$agentDataDir\log"
$daemonLogDir="$agentInstDir\bin\daemon"
$domotzService = "Domotz Agent"
$domotzNode="$agentInstDir\bin\domotz_node.exe"
$currentDir=$PSScriptRoot

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

$customerLogDir="agent-Logs-$date"
$reportFile="$DesktopPath\$customerLogDir\agent_short_report.txt"

# Creating Log dir
if (!(Test-Path $DesktopPath\$customerLogDir -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir | Out-Null
}
Add-Content $reportFile "--Agent Short Report $date--"
Add-Content $reportFile ""
Add-Content $reportFile "[Agent Details]"

# Get Domotz Agent service properties 
$domotzServiceFile="$DesktopPath\$customerLogDir\DomotzServiceProperties.txt"
Get-Service $domotzService | Select-Object Displayname,Status | Out-File $domotzServiceFile

# Get Computer info
$osinfoFile="$DesktopPath\$customerLogDir\os_info.txt"
Get-ComputerInfo | Out-File $osinfoFile 

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

# check for Npcap issue - This has to be reviewed - have to find better evidence $match_str is too generic #TODO
Write-Host ""
Write-Host -noNewLine "-> Checking for win Npcap issues please wait..."

$npcap_info=Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate | Where-Object -Property DisplayName -Match Npcap
$npcap_version=$npcap_info.DisplayVersion
Add-Content $reportFile "Npcap version=$npcap_version"


$match_str="Cannot find MAC address for device with IP"
if(Select-String -Path $DesktopPath\$customerLogDir\listener_logs\*.log.* -Patter $match_str){
    $npcapIssueReport="$DesktopPath\$customerLogDir\npcap_issue_maybe_detected.txt"
    Add-Content $npcapIssueReport "### NB: This has to be reviewed and could not be accurate! have to find better evidence $match_str is too generic!!"
    Select-String -Path $DesktopPath\$customerLogDir\listener_logs\*.log.* -Patter $match_str | Out-File -Encoding Ascii -Append $npcapIssueReport
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
Write-Host "Domotz Diagnostics has finished his job! Thank you for using it!"
Write-Host ""
Write-Host "Please, press ENTER to EXIT or close this window!" -NoNewLine
$UserInput = $Host.UI.ReadLine()