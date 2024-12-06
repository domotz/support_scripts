Write-Host ""
Write-Host -noNewLine "-> Collecting Domotz agent process properties...."

    $winFwEnabled=Get-NetFirewallProfile | Select-Object Name -Expandproperty Enabled
    if ($winFwEnabled -contains "True") { 
        Add-Content $warningsFile ""
        Add-Content $warningsFile "-> WARNING: Windows Firewall is enabled please check the windows_firewall.txt and windows_firewall_rules.txt for more info"
        Write-Host "True"
      }