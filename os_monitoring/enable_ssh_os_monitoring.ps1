# Domotz script to enable SSH for OS Monitoring
$dscriptver="0.1"

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
Write-Output "This utility will enable SSH on your Windows OS to unlock the Domotz OS Monitoring feature.  (ver. $dscriptver)
"

$openSSHClientState=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Client*' | Select-Object -ExpandProperty "State"
$openSSHServerState=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*' | Select-Object -ExpandProperty "State"
$openSSHClientVer=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Client*' | Select-Object -ExpandProperty "Name"
$openSSHServerVer=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*' | Select-Object -ExpandProperty "Name"

if (!($openSSHClientState -eq "Installed")){
    Write-Output "OpenSSH Client is not installed on this system... we are going to install it"
    Read-Host -Prompt "-> please press ENTER to install it or CTRL+C to quit"
    Write-Output "
    OpenSSH Client is installing... please wait....."
    Add-WindowsCapability -Online -Name $openSSHClientVer
    Write-Output "Done! -> OpenSSH Client is installed!
    "
}
else {
    Write-Output "-> OpenSSH Client is installed!
    "
}

if (!($openSSHServerState -eq "Installed")){
    Write-Output "OpenSSH Server is not installed on this system... we are going to install it"
    Read-Host -Prompt "-> please press ENTER to install it or CTRL+C to quit"
    Add-WindowsCapability -Online -Name $openSSHServerVer
    Write-Output "
    OpenSSH Server is installing... please wait.....
    "
    Write-Output "Done! -> OpenSSH Server is installed!"
}
else {
    Write-Output "-> OpenSSH Server is installed!
    "
}

# Starting sshd
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Checking Firewall Rule for SSH inbound
$sshdFirewallInboundRuleName=Get-NetFirewallRule -Name *ssh* |  Select-Object -ExpandProperty "Name" #should be OpenSSH-Server-In-TCP

if (!(($sshdFirewallInboundRuleName -eq "OpenSSH-Server-In-TCP") -or ($sshdFirewallInboundRuleName -eq "OpenSSH Server (sshd)"))) {
    Write-Output "
    No Firewall sshd inbound rule"
    Read-Host -Prompt "-> please press ENTER to create one or CTRL+C to quit"
    New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
    Write-Output "-> OpenSSH Server (sshd) Firewall Inbound Rule created!
    "
}
else {
    Write-Output "-> OpenSSH Server (sshd) Firewall Inbound Rule is present!
    "
}

$sshdFirewallInboundRuleStatus=Get-NetFirewallRule -Name *ssh* |  Select-Object -ExpandProperty "PrimaryStatus" #should be OK

if (!($sshdFirewallInboundRuleStatus -eq "Ok")){
    Write-Output "
    Standard SSHd Firewall Rule seems to be in a wrong status"
    Read-Host -Prompt "-> please press ENTER to create a new one or CTRL+C to quit"
    New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
    Write-Output "-> Another OpenSSH Server (sshd) Firewall Inbound Rule created (the other one had the wrong state!)
    "
}

if (Test-Path -Path $currentDir\Set-WmiNamespaceSecurity.ps1 -PathType Leaf) {
    Read-Host -Prompt "IMPORTANT-> Do you want to use a non administrative user to unlock your Windows machine?
    IF you do, please press ENTER, otherwise hit CTRL+C to QUIT!"
    Write-Output ""

    do {
        $username=Read-Host -Prompt "Please enter the username you want to use"
        $response = Read-Host "Are you sure the user is " "[$username]" "? (Y|N)"
    }
    until (($response -eq "y") -or ($response -eq "Y"))

    $currentDir=$PSScriptRoot
    &"$currentDir/Set-WmiNamespaceSecurity.ps1" root/cimv2 add $username Enable,RemoteAccess
    &"$currentDir/Set-WmiNamespaceSecurity.ps1" root/cimv2/Security/MicrosoftVolumeEncryption add $username Enable,RemoteAccess

    Write-Output "-> Added permissions to WMI for user [$username]

    ##### SSH Configuration for OS_monitoring is completed!!

    IMPORTANT: It might take up to 3 hours to see the unlock option in the Domotz Access Management, depending on the size of your network.
    "
}
else {
    Write-Output "###### SSH Configuration for OS_monitoring is completed for Administrative Users!


PLEASE NOTE THAT: It might take up to 3 hours to see the unlock option in the Domotz Access Management, depending on the size of your network.


####################################################################################################################
If you want to enable the unlock also for non-administrative users, please download the Set-WmiNamespaceSecurity.zip 
from https://github.com/domotz/support_scripts/windows_agent and check the README.txt file before using it
####################################################################################################################
    
    "
}
