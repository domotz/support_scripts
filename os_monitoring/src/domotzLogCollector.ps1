<#
.SYNOPSIS
The Script collects information on the system it runs on

.DESCRIPTION
The script collects the following information:
- Environment variables
- WinRM configuration
- WinRM permissions
- WMI Permissions
- Log file of the last execution of "enable_winrm_os_monitoring.ps1"

.PARAMETER LogfilePath
Path of the log files for this script

.PARAMETER WorkPath
Path where the zip file will be created

#>

param (

    [Parameter()]
    [string]$LogfilePath = $PSScriptRoot,
    [string]$WorkPath = $PSScriptRoot
)

if (!(Test-Path $WorkPath)) {
    New-Item -ItemType Directory -Path $WorkPath -Force | Out-Null
}
$PSVer = ($PSVersionTable.PSVersion.major)
function wo {
    param(
        $msg
    )
    $Spacer = " -- "
    $TimeStamp = $(Get-Date -Format "yyyy-MM-dd_hh:mm:ss")    
    Write-Host $($TimeStamp + $Spacer + $msg)
}
function __HumanizeAccessMask($decimalValue) {
    # Define the permission values
    $permissions = @{
        0x1     = "WBEM_ENABLE (Grants read permissions)"
        0x2     = "WBEM_METHOD_EXECUTE (Grants execute methods)"
        0x4     = "WBEM_FULL_WRITE (Grants write to classes and instances)"
        0x8     = "WBEM_PARTIAL_WRITE_REP (Grants update or delete static CIM instances)"
        0x10    = "WBEM_WRITE_PROVIDER (Grants update or delete dynamic CIM instances)"
        0x20    = "WBEM_REMOTE_ENABLE (Grants remote access to the server)"
        0x20000 = "READ_CONTROL (Allows reading the security descriptor of CIM namespace)"
        0x40000 = "WRITE_DAC (Allows modifying the security descriptor of CIM namespace)"
    }

    [System.Collections.ArrayList]$translatedPermissions = @()

    foreach ($key in $permissions.Keys) {
        if ($decimalValue -band $key) {
            $translatedPermissions += $permissions[$key]
        }
    }

    $translatedPermissions
}

function Get-WbemPermissions {
    $AceTypes = @{
        0 = "Access_Allowed" 
        1 = "Access_Denied"
        2 = "Audit"
    }

    $sd = Invoke-WmiMethod -Namespace root\cimv2 -path "__systemsecurity=@" -name GetSecurityDescriptor
    $dacl = $sd.descriptor | Select-Object -ExpandProperty dacl
    $AccessData = @{}

    foreach ($d in $dacl) {
        $trustee = ($d.trustee).Domain + '\' + ($d.trustee).Name
        $permissions = ( __HumanizeAccessMask($($d.AccessMask))) -join ";"
        $accessType = $AceTypes[[int]$d.AceType]

        if ($AccessData.ContainsKey($trustee)) {
            # Update existing entry
            $existingPermissions = $AccessData[$trustee].Permissions
            if (-not $existingPermissions.Contains($permissions)) {
                $AccessData[$trustee].Permissions += ";$permissions"
            }
        } else {
            # Add new entry
            $AccessData[$trustee] = [PSCustomObject] @{
                trustee     = $trustee
                Permissions = $permissions
                AccessType  = $accessType
            }
        }
    }

    return $AccessData.Values
}


$dscriptver = "2"
try {
    Stop-Transcript -ErrorAction Stop | Out-Null
}
catch {}
$LogFile = "$LogFilePath\$ENV:COMPUTERNAME-$($($MyInvocation.MyCommand.Name).replace('.ps1',''))-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$RootSDDL = Get-Item WSMan:\localhost\Service\RootSDDL | Select-Object -ExpandProperty Value
$RegKeyPath = "HKLM:\Software\DomotzScripting\enableWinRm"
$RegPropertyName = "LastLogFile"
Start-Transcript -Path $LogFile
$VersionBanner = "$($MyInvocation.MyCommand.Name) Version $dscriptver"
$line = (0..$($VersionBanner.length / 2 )) | ForEach-Object { $line + "-" }
Write-Host $line -ForegroundColor White -BackgroundColor Blue
Write-Host "$VersionBanner " -ForegroundColor White -BackgroundColor Blue 
Write-Host $line -ForegroundColor White -BackgroundColor Blue
Write-Host ''
wo "LogFile for this script is $LogFile"
try {
    wo "Environment variables ------------------------------------------"
    Get-ChildItem ENV: | Where-Object { ($_.Name) -NotMatch "DOMOTZ_USER_PASS" }
    wo "WinRm service configuration ------------------------------------------"
    wo "   localhost"
    Get-ChildItem WSMan:\localhost -ErrorAction Stop | Format-Table  -AutoSize -Wrap 
    wo "   Service"
    Get-ChildItem WSMan:\localhost\Service -ErrorAction Stop | Format-Table  -AutoSize -Wrap
    wo "   Auth"
    Get-ChildItem WSMan:\localhost\Service\Auth -ErrorAction Stop | Format-Table  -AutoSize -Wrap
    wo "   Listener"
    Get-ChildItem WSMan:\localhost\Listener -ErrorAction Stop | Select-Object -ExpandProperty Keys
    wo "   Shell"
    Get-ChildItem WSMan:\localhost\Shell -ErrorAction Stop | Format-Table  -AutoSize -Wrap
}
catch {
    wo "Error getting ENV or WSMan items: $_"
}

wo "WinRm service permissions ------------------------------------------"
if ([int]$PSVer -ge 5) {
    Write-Host $(ConvertFrom-SddlString $RootSDDL).DiscretionaryAcl
}
else {
    wo "Cannot translate sddl, PS version not suported, logging it raw"
    wo "RootSDDL:"
    (Get-ChildItem WSMan:\localhost\Service\RootSDDL).Value
}
wo "WMI service permissions ------------------------------------------"
Get-WbemPermissions | Format-List
wo " Trying to collect script log file "
wo "  == Reading property $RegPropertyName in $RegKeyPath "
try {
    $enableWinRmLogfile = Get-ItemProperty -Path $RegKeyPath -Name $RegPropertyName -ErrorAction Stop | Select-Object -ExpandProperty LastLogFile
    wo "  == enableWinRm log file should be in $enableWinRmLogfile"
}
catch {
    wo "  == $RegKeyPath or $RegPropertyName property not found, make sure you ran the enableWinRm script version 0.7 or greater"
}

Add-Type -assembly "system.io.compression.filesystem"
$TimeStamp = $(Get-Date -Format 'yyyy-MM-dd_HH-MM-ss')
$LogDir = "$ENV:TEMP\Domotz-$Timestamp"
$ZipFile = "$WorkPath\DomotzWinrmLog-$Timestamp.zip"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
wo "LogDir: $LogDir"
wo "enableWinRM script log: $enableWinRmLogfile"
wo "current log: $Logfile"
Stop-Transcript
Copy-Item -Path $Logfile -Destination $LogDir
if (![string]::IsNullOrEmpty($enableWinRmLogfile)) {
    if (Test-Path $enableWinRmLogfile) {
        wo "Adding $enableWinRmLogfile to log folder to compress"
        Copy-Item -Path $enableWinRmLogfile -Destination $LogDir
    }
}
else {
    wo "Log file of enable_winrm_os_monitoring.ps1 not found, make sure you ran the enableWinRm script version 0.7 or greater "
}
wo "Creating $ZipFile"
[io.compression.zipfile]::CreateFromDirectory($LogDir, $ZipFile)
Start-Process "C:\Windows\explorer.exe" -ArgumentList $WorkPath
