# Domotz Agent -- Uninistall script ---
# 

Write-host "
Domotz uninistall script
This can take a while... 
"
$domotz_agent = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall  |
    Get-ItemProperty |
        Where-Object {$_.DisplayName -match "Domotz Agent" } |
            Select-Object -Property DisplayName, UninstallString

ForEach ($ver in $domotz_agent) {

    If ($ver.UninstallString) {

        $uninst = $ver.UninstallString
        & cmd /c $uninst /S
    }
}