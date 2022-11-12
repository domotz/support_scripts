# Domotz Agent -- Powershell script (this is just an example you can customize it the way you want)

# Please edit the following lines 
$ActivationKey = "You API key here"
$ApiEndpoint = "Your API endpoint here, like https://api-eu-west-1-cell-1.domotz.com/public-api/v1/ or https://api-us-east-1-cell-1.domotz.com/public-api/v1/"
$AgentName = "Your Agent name here"

$AgentUrl = "https://portal.domotz.com/download/agent_packages/domotz-windows-x64-10.exe"

$WindowsAgentInstallerFile = "$PSScriptRoot\domotz-windows-x64-10.exe"
$WindowsAgentInstallerDir = $PSScriptRoot
$StatusUrl = "http://127.0.0.1:3000/api/v1/status"
$ActivationUrl = "http://127.0.0.1:3000/api/v1/agent"
$ActivationHeaders = @{
    "X-API-Key" = $ActivationKey
}
$ActivationBody = @{
    "name" = $AgentName
    "endpoint" = $ApiEndpoint

} | ConvertTo-Json

# Check if you have administrative privileges to run this script
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}
else {
    Write-Host "Code is running as administrator - nice to hear that!" -ForegroundColor Green
}

# Download Domotz Agent Installer
Write-Host "Downloading agent from $AgentUrl"
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($AgentUrl, $WindowsAgentInstallerFile)

# Installing the Domotz Agent
Write-Host "Executiing installer $WindowsAgentInstallerFile working dir=$WindowsAgentInstallerDir"
Start-Process -FilePath $WindowsAgentInstallerFile -WorkingDirectory $WindowsAgentInstallerDir -ArgumentList "/W /S /D=`"C:\Program Files (x86)\domotz\`""

$IsInstalled = $false
do {
    Get-Process domotzagent -ErrorAction SilentlyContinue -ErrorVariable ProcessError 
    if ($ProcessError) {
        Write-Host "Please wait for installation to finish, this might take a while (this message may repeat until finished)..."
        Start-Sleep -s 2
    }
    else {
        Write-Host "Installation completed!"
        $IsInstalled = $true
    }
}
while ($IsInstalled -eq $false)

$IsRunning = $false
do {
    try {
        $StatusResponse = Invoke-RestMethod -Uri $StatusUrl
        $IsRunning = $true
        Write-Host "Agent is running. Proceeding with activation."
    } catch {
        Write-Host "Waiting for Agent to start, please wait..."
        Start-Sleep -s 2
    }
}
while ($IsRunning -eq $false)

# Activate Agent with name provided in $AgentName
try {
    Invoke-RestMethod -Uri $ActivationUrl -Method Post -Headers $ActivationHeaders -ContentType "application/json" -Body $ActivationBody
    Write-Host "Agent activated"
} catch {
    Write-Host "Something went wrong during agent activation"
    Write-Host "Status Code: " $_.Exception.Response.StatusCode.value__
    $BodyError = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
    $BodyError.BaseStream.Position = 0
    $BodyError.ReadToEnd()
}