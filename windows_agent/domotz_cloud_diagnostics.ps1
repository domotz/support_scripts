# Domotz script to check connection to the Domotz Cloud and default dns for Windows

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
Write-Host "This script check that the connection to the Domotz Cloud is reliable (ver. $dscriptver)
"

function Get-Zone
{
	$area=Read-Host "Choose a site code"
	Switch ($area)
	{
		1 {$choice="us"}
		2 {$choice="eu"}
		3 {$choice="apac"}
        default {
            Write-Host "Wrong option chosen"
            Get-Zone
        }
		}
    return $choice
}

Write-Host "
1. USA
2. EU
3. APAC
"
$cell=Get-Zone

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

Write-Host "Loading..."
Write-Host ""

if ($cell -eq "us") {
    $hosts=$ushosts
}
if ($cell -eq "eu" -or $cell -eq "apac") {
    $hosts=$euhosts
}


Write-Host -noNewLine "-> Testing network connection to Domotz Cloud... please wait...

"
$openonfw="$date - [E] - KO -- Connection Error - Please open on your Firewall for OUTGOING connections to:"

if (Test-Connection -ComputerName $echoHost -Quiet) { 
    Write-Host "Ping to echo.domotz.com succeded
    "
    
}
else {
    Write-Host "$date - [E] - KO -- Ping to echo.domotz.com unsuccessful
    "
}
if ($dnsServers) {
    Write-Host  "DNS settings OK
    "
}
else {
    Write-Host "[W] - WARNING - Please use Google Public DNS for your Domotz agent host machine! (8.8.8.8 and 8.8.4.4)
    "
}

foreach ($a in $hosts) {
    $null = (Test-NetConnection -ComputerName $a.host -Port $a.port -ErrorAction SilentlyContinue -ErrorVariable ConnectionError).TcpTestSucceeded
    
    $ahost=$a.host
    $aport=$a.port
    $aregion=$a.region

    if ($ConnectionError) {
        Write-Host "[W] This is required by Region: $aregion"
        Write-Host "[E] - KO -- $openonfw $ahost on Port $aport"
        Write-Host ""
    }
    else {
        Write-Host "Connection to $ahost - $aport OK"
        Write-Host ""
    }
}

Write-Host "N.B. To remotely connect to your devices  please make sure that the following host/port-range is allowed on your firewall:"
if ($cell -eq "eu") {
    Write-Host "sshg.domotz.co (range: 32700 - 57699 TCP)
    "
}
if ($cell -eq "apac") {
    Write-Host "ap-southeast-2-sshg.domotz.co(range: 32700 - 57699 TCP)
    "
}
if ($cell -eq "us") {
    Write-Host "us-east-1-sshg.domotz.co, us-east-1-02-sshg.domotz.co and us-west-2-sshg.domotz.co (range: 32700 - 57699 TCP)
    "
}