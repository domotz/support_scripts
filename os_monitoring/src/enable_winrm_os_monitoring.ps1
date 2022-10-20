# Domotz script to enable WINRM on Microsoft Windows for OS Monitoring
# Please read the README file before using

$dscriptver="0.1.2"
$currentDir=$PSScriptRoot
$errorFile="./error.log"

# Check if you have administrative privileges to run this script
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this utility. You should run this with administrative privileges."
    break
}
else {
    Write-Information "Code is running as administrator - nice to hear that!"
}


# Motd
Write-Output "
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
Write-Output "This utility will enable WINRM on Microsoft Windows to unlock the Domotz OS Monitoring feature.  (ver. $dscriptver)
"

# Chck for 
$networkCategory = Get-NetConnectionProfile | Select-Object -ExpandProperty NetworkCategory
if ($networkCategory -inotmatch "Private") {
    $response = Read-Host "WinRM firewall exception will not work since one of the network connection types on this machine is set to Public. 
    Do yu want to set it to Private? [Y,n]"
    if ($response -eq 'y' -Or $response -eq "Y"){
        Set-NetConnectionProfile -NetworkCategory Private
    }
    else {
        break
    }
}

# Setting up the WinRM configuration
Write-Output "-> Setting up WINRM service...
    "
winrm quickconfig 
  
winrm set winrm/config/service/Auth '@{Basic="true"}'
winrm set winrm/config/service '@{AllowUnencrypted="true"}'
winrm set winrm/config/winrs '@{MaxMemoryPerShellMB="1024"}'

# Testing WinRM configuration
winrm set winrm/config/client '@{AllowUnencrypted="true"}'

# $host_ip=Read-Host -Prompt "Enter actual IP of the Windows machine you want to test/monitor"
# winrm set winrm/config/client '@{TrustedHosts="' + $host_ip +'"}'
$host_ip = "127.0.0.1"
$wmiuser=Read-Host -Prompt "Enter your winrm your username to test"
$response = Read-Host "IMPORTANT-> Is the user you have entered a non administrative user? [Y,n]"
    if ($response -eq 'y' -Or $response -eq "Y"){
        (Get-PSSessionConfiguration -Name Microsoft.PowerShell).Permission
        Net localgroup "Remote Management Users" /add $wmiuser
    }
    else {
        break
    }

$Credential = Get-Credential $wmiuser 

Try {
    Write-Host "
    Testing if the WinRM service is correctly configured for Domotz---
    "

    Read-Host "
    The response you will get from testion the service should be very similar to the one belows:

    wsmid           : http://schemas.dmtf.org/wbem/wsman/identity/1/wsmanidentity.xsd
    ProtocolVersion : http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd
    ProductVendor   : Microsoft Corporation
    ProductVersion  : OS: 10.0.19042 SP: 0.0 Stack: 3.0

    When ready press ENTER
    "

    Write-Host "

    ######WIRM RESPONSE#####:
    "
    
    test-wsman $host_ip -Authentication Basic -Credential $Credential -ErrorAction stop

}
Catch {
    Write-host "
    WARNING!!!
    There was en error in testing your WinRm configuration, please check the contents of the $errorFile file."
    $_.Exception.Message *> $errorFile
}
