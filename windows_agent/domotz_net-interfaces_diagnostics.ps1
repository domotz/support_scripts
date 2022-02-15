# Domotz script to check and report on Network Interfaces

$dscriptver="0.1"

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
Write-Host "This script checks Windows Host OS and Network Interfaces ver. $dscriptver)
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

$customerLogDir="domotz-OS-Nics-info-logs-$date"
$reportFile="$DesktopPath\$customerLogDir\agent_short_report.txt"

# Creating Log dir
if (!(Test-Path $DesktopPath\$customerLogDir -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $DesktopPath\$customerLogDir | Out-Null
}
Add-Content $reportFile "--Agent Short Report $date--"
Add-Content $reportFile ""
Add-Content $reportFile "[Agent Details]"

# Get Computer info
$osinfoFile="$DesktopPath\$customerLogDir\os_info.txt"
Get-ComputerInfo | Out-File $osinfoFile 

# Get Interfaces Long info
$longNicInfoFile="$DesktopPath\$customerLogDir\interfaces_long_info.txt"
Get-NetAdapter -Name * | Format-List -Property * | Out-File $longNicInfoFile

# Get Interfaces Brief info
$briefNicInfoFile="$DesktopPath\$customerLogDir\interfaces_brief_info.txt"
Get-NetAdapter -Name * | Out-File $briefNicInfoFile

# Get Interfaces IP Info
$ipNicInfoFile="$DesktopPath\$customerLogDir\interfaces_brief_info.txt"
Get-NetIPConfiguration -All | Out-File $ipNicInfoFile

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
$flushLog=Invoke-WebRequest -URI http://127.0.0.1:3000/api/v1/log/flush
if (!$flushLog) {
    Add-Content $reportFile "WARNING: Unable to flush logs from 127.0.0.1:3000"
}
else {
    $flushLog | ConvertFrom-Json | Out-File $DesktopPath\$customerLogDir\listener_logs\flushed_log.txt
}
Write-Host " Done!"


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