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
    Write-Host "Code is running as administrator - nice to hear that!" -ForegroundColor Green
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
| The RMM tool for Networks and Connected Devices|
+------------------------------------------------+
"
Write-Host "This utility will enable SSH on your Windows OS to unlock the Domotz OS Monitoring feature.  (ver. $dscriptver)
"

$openSSHClientState=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Client*' | Select-Object -ExpandProperty "State"
$openSSHServerState=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*' | Select-Object -ExpandProperty "State"
$openSSHClientVer=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Client*' | Select-Object -ExpandProperty "Name"
$openSSHServerVer=Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*' | Select-Object -ExpandProperty "Name"

if (!($openSSHClientState -eq "Installed")){
    Write-Host "OpenSSH Client is not installed on this system... we are going to install it"
    Read-Host -Prompt "-> please press ENTER to install it or CTRL+C to quit" 
    Write-Host "
    OpenSSH Client is installing... please wait....."
    Add-WindowsCapability -Online -Name $openSSHClientVer
    Write-Host "Done! -> OpenSSH Client is installed!
    "   
}
else {
    Write-Host "-> OpenSSH Client is installed!
    "
}

if (!($openSSHServerState -eq "Installed")){
    Write-Host "OpenSSH Server is not installed on this system... we are going to install it"
    Read-Host -Prompt "-> please press ENTER to install it or CTRL+C to quit"
    Add-WindowsCapability -Online -Name $openSSHServerVer 
    Write-Host "
    OpenSSH Server is installing... please wait.....
    "
    Write-Host "Done! -> OpenSSH Server is installed!"
}
else {
    Write-Host "-> OpenSSH Server is installed!
    "
}

# Starting sshd
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Checking Firewall Rule for SSH inbound
$sshdFirewallInboundRuleName=Get-NetFirewallRule -Name *ssh* |  Select-Object -ExpandProperty "Name" #should be OpenSSH-Server-In-TCP

if (!(($sshdFirewallInboundRuleName -eq "OpenSSH-Server-In-TCP") -or ($sshdFirewallInboundRuleName -eq "OpenSSH Server (sshd)"))) {
    Write-Host "
    No Firewall sshd inbound rule"
    Read-Host -Prompt "-> please press ENTER to create one or CTRL+C to quit"
    New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
    Write-Host "-> OpenSSH Server (sshd) Firewall Inbound Rule created!
    "
}
else {
    Write-Host "-> OpenSSH Server (sshd) Firewall Inbound Rule is present!
    "
}

$sshdFirewallInboundRuleStatus=Get-NetFirewallRule -Name *ssh* |  Select-Object -ExpandProperty "PrimaryStatus" #should be OK

if (!($sshdFirewallInboundRuleStatus -eq "Ok")){
    Write-Host "
    Standard SSHd Firewall Rule seems to be in a wrong status"
    Read-Host -Prompt "-> please press ENTER to create a new one or CTRL+C to quit"
    New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
    Write-Host "-> Another OpenSSH Server (sshd) Firewall Inbound Rule created (the other one had the wrong state!)
    "
}

Read-Host -Prompt "IMPORTANT-> Will you be using a non administrative user to unlock your Windows machine?
IF you will please press ENTER, otherwise hit  CTRL+C to QUIT!"
Write-Host ""

do {
    $username=Read-Host -Prompt "Please enter the username you want to use"
    $response = Read-Host "Are you sure the user is " "[$username]" "? (Y|N)"
}
until (($response -eq "y") -or ($response -eq "Y"))

$currentDir=$PSScriptRoot
&"$currentDir/Set-WmiNamespaceSecurity.ps1" root/cimv2 add $username Enable,RemoteAccess
