<#
.SYNOPSIS
    Domotz script to enable WINRM on Microsoft Windows for OS Monitoring

.DESCRIPTION
    The scripts grants read, execute WinRM permissions to a group (AD or local).
    When providing domain user and group, the domain must be specified (DOMAIN\User , DOMAIN\Group)
    - When the group is AD:
        . It must exist
        . The provided domain user must be a member of the group, if not the script terminates
    - When the group is local
        . It will be created if it doesn't exist
        . If the user is a domain one it must exist
        . If the user is local it will be created if it doesn't exist

.PARAMETER UserName
    The user member of the WinRM group, if it's domain one the domain must be specified in NT form (DOMAIN\UserName)

.PARAMETER Pass
    This is the password to assign to a newly created user.
    By default it tries to get the value from the environment variable DOMOTZ_USER_PASS.
    If the value is empty and the local user does not exists a random password is generated.

.PARAMETER GroupName
    The group that will be granted WinRM permissions, if it's domain one the domain must be specified in NT form (DOMAIN\GroupName)

.PARAMETER NetworkProfile
    The script won't work if the network profile is set to Public.
    The value for this paramenter can be "Domain" or "Private", the default is "Private".
    Warning!! the script will change the profile to all the NICs that are set as 'Public'.

.PARAMETER WmiAccessOnly
    Only adds WMI permissions to specified WMI namespaces

.PARAMETER Namespaces
    Can only be used with WmiAccessOnly parameter and specifies the list of WMI namespace(s) 
    to add the permissions to, default  @("root\cimv2", "Root\Microsoft\Windows\Storage")

.EXAMPLE

        .\enable_winrm_os_monitoring_new.ps1 -UserName domotz\domotztestuser -GroupName domotz\ddomaingrp
        Checks if the group exists in AD and the user is a member of the group, if not it terminates (no attempt to create objects in AD are made by the script).
        If the group exists and the user is a member of the group the script grants permissions to the group on the WinRM default listener
.EXAMPLE
   
        .\enable_winrm_os_monitoring_new.ps1 -UserName domotzlocaluser -GroupName domotz\ddomaingrp
        Since the group is a domain one, the script assumes the user is in the same domain and a member of the group
.EXAMPLE
    
        .\enable_winrm_os_monitoring_new.ps1 -UserName domotzlocaluser -GroupName domotzLocalGroup
        Group and user will be created locally if missing, the user will be added to the group if not there already and permissions will be granted to the group
.EXAMPLE
    
        .\enable_winrm_os_monitoring_new.ps1 -UserName domotz\domotztestuser -GroupName ddomaingrp
        checks if the uesr exists in AD
        checks if the group exists locally since no domain is provided
        add the user to the local group and grant the group permissions on the default WinRM listener

.EXAMPLE

        .\enable_winrm_os_monitoring.ps1 -GroupName adlab\domotzwinrm -WmiAccessOnly -Namespaces "Root\Microsoft\Windows\Storage"
        Add WMI permissions to adlab\domotzwinrm group on namespace "Root\Microsoft\Windows\Storage"
.NOTES
        If using a local user that does not exist, the options to assign it a password on creation are the environment variable DOMOTZ_USER_PASS
        or the commandline parameter 'Pass'. If none of those are present the password will be random and revealed on the console.
        The best option is to create the user before running the script so the password will not travel and will be known.
#>
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = "AllOperations")]
param (
    [Parameter(ParameterSetName = 'AllOperations', Mandatory = $true)]
    [string]$UserName,

    [Parameter(ParameterSetName = 'AllOperations')]
    [string]
    $Pass = $ENV:DOMOTZ_USER_PASS,

    [Parameter(ParameterSetName = 'AllOperations', Mandatory = $true)]
    [Parameter(ParameterSetName = 'WmiOnly', Mandatory = $true)]
    [string]$GroupName = "DomotzWinRM",

    [Parameter(ParameterSetName = 'AllOperations')]
    [Parameter(ParameterSetName = 'WmiOnly')]
    [string]$LogfilePath = "$PSScriptRoot",

    [Parameter(ParameterSetName = 'AllOperations')]
    [ValidateSet("Domain", "Private")]
    [string]$NetworkProfile = "Private",
	
    [Parameter(ParameterSetName = 'WmiOnly')]
    [switch]$WmiAccessOnly,
	
    [Parameter(ParameterSetName = 'WmiOnly', Mandatory = $true)]
    [string[]]$Namespaces = @("root\cimv2", "Root\Microsoft\Windows\Storage")
)



# Define the registry key path and value name
function Write-LogLocation {
    param (
        [string]$LogFileFullPath
    )
    $regKeyPath = "HKLM:\Software\DomotzScripting\enableWinRm"
    $PropertyName = "LastLogFile"

    if (-not (Test-Path -Path $regKeyPath)) {
        New-Item -Path $regKeyPath -Force
    }

    $ItemValue = Get-ItemProperty -Path $regKeyPath -Name $PropertyName -ErrorAction SilentlyContinue

    if ($ItemValue) {

        Set-ItemProperty -Path $regKeyPath -Name $PropertyName -Value $LogFileFullPath
    }
    else {

        New-ItemProperty -Path $regKeyPath -Name $PropertyName -Value $LogFileFullPath -PropertyType String | Out-Null
    }

    
}

function Set-WMIAcl {
    # Excerpt from https://github.com/grbray/PowerShell/blob/main/Windows/Set-WMINameSpaceSecurity.ps1
    
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Account,
        [Parameter()]
        [string]$Namespace

    )

    $CONTAINER_INHERIT_ACE_FLAG = 0x2
    $ACCESS_ALLOWED_ACE_TYPE = 0x0
    $WBEM_METHOD_EXECUTE = 0x02
    $WBEM_REMOTE_ACCESS = 0x20


    $InvokeParams = @{Namespace = $Namespace; Path = '__systemsecurity=@'; ComputerName = $ENV:ComputerName }
    Write-Output "GetSecurityDescriptor Parameters:"
    Write-Output "Namespace = $Namespace"
    Write-Output "Path = __systemsecurity=@"
    Write-Output "ComputerName = $ENV:ComputerName"
    
    
    try {
        $output = Invoke-WmiMethod @InvokeParams -Name "GetSecurityDescriptor" -ErrorAction Stop
    }
    catch {
        Write-Error "Invoke-WmiMethod GetSecurityDescriptor failed: $_"
        Write-Host "Make sure there are not unresolved/unknown SIDs in the ACL of WMI security" -ForegroundColor Red -BackgroundColor White
        $output
        return $false
    }

    $ACL = $output.Descriptor

    if ($Account.Contains('\')) {
        $Domain = $Account.Split('\')[0]
        if (($Domain -eq ".") -or ($Domain -eq "BUILTIN")) { $Domain = $ENV:ComputerName }
        $AccountName = $Account.Split('\')[1]
    }
    elseif ($Account.Contains('@')) {
        $Domain = $Account.Split('@')[1].Split('.')[0]
        $AccountName = $Account.Split('@')[0]
    }
    else {
        $Domain = $ENV:ComputerName
        $AccountName = $Account
    }

    $GetParams = @{Class = "Win32_Account" ; Filter = "Domain='$Domain' and Name='$AccountName'" }
    $Win32Account = Get-WmiObject @GetParams
    if ($null -eq $Win32Account) { throw "Account was not found: $Account" }
   
    $ACE = (New-Object System.Management.ManagementClass("Win32_Ace")).CreateInstance()
    $ACE.AccessMask = $WBEM_METHOD_EXECUTE + $WBEM_REMOTE_ACCESS
    # Do not use $OBJECT_INHERIT_ACE_FLAG.  There are no leaf objects here.
    $ACE.AceFlags = $CONTAINER_INHERIT_ACE_FLAG 
 

    $Trustee = (New-Object System.Management.ManagementClass("Win32_Trustee")).CreateInstance()
    $Trustee.SidString = $Win32Account.SID
    $ACE.Trustee = $Trustee

    $ACE.AceType = $ACCESS_ALLOWED_ACE_TYPE
    $ACL.DACL += $ACE
    $SetParams = @{Name = "SetSecurityDescriptor"; ArgumentList = $ACL } + $InvokeParams
    $SetParams

    try {
        $output = Invoke-WmiMethod @SetParams -ErrorAction Stop
    }
    catch {
        Write-Error "Invoke-WmiMethod SetSecurityDescriptor failed: $_"
        $output
        return $false
    }
    
    return $true

}
function Add-WinRMDaclRule {
    <#
    .SYNOPSIS
        Add a Discretionary Acl rule to the root WinRM listener.
    .NOTE
        This function is an excerpt of the one that can be found at https://gist.github.com/jborean93/6d9aaf868d1d40344188984ebb431b04
        You may need to restart the WinRM service for these changes to apply, run 'Restart-Service -Name winrm' to do so.
        If you wish to just enable a standard user account access over PSRemoting, you can also just add it to the builtin
        'Remote Management Users' group on the host in question.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param (

        [Parameter(Mandatory = $true)]
        [string]
        $Account,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [System.String[]]
        $Right = @('Read', 'Execute')
    )

    Begin {

        Write-Verbose -Message "Getting Root WSMan SDDL"
        $sddl = (Get-Item -LiteralPath WSMan:\localhost\Service\RootSDDL).Value

        $sd = New-Object -TypeName 'System.Security.AccessControl.CommonSecurityDescriptor' -ArgumentList @(
            $false, $false, $sddl
        )
        $accessMask = @{
            FullControl = 0x10000000
            Execute     = 0x20000000
            Write       = 0x40000000
            Read        = 0x80000000
        }
    }

    Process {
        Write-Verbose -Message "Validating the input rights"
        $mask = 0
        foreach ($aceRight in $Right) {
            if (-not $accessMask.ContainsKey($aceRight)) {
                Write-Error -Message "Invalid access right '$aceRight' - skipping this entry, valid values are: $($accessMask.Keys)."
                return
            }
            $mask = $mask -bor $accessMask.$aceRight
        }
 
        Write-Verbose -Message "Converting '$userAccount' to a Security Identifier"
        $sid = (New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $Account).Translate(
            [System.Security.Principal.SecurityIdentifier]
        )

        $addRule = $true
        foreach ($ace in $sd.DiscretionaryAcl.GetEnumerator()) {
            if ($ace.SecurityIdentifier -ne $sid) {
                continue
            }
            if ($ace.AceType -ne 'AccessAllowed') {
                continue
            }
            if ($ace.AccessMask -ne $mask) {
                continue
            }

            $addRule = $false
            break
        }

        if ($addRule) {
            Write-Verbose -Message "Adding rule for $userAccount with rights $($Rights -join ", ")"
            $sd.DiscretionaryAcl.AddAccess(
                [System.Security.AccessControl.AccessControlType]::Allow,
                $sid,
                $mask,
                [System.Security.AccessControl.InheritanceFlags]::None,
                [System.Security.AccessControl.PropagationFlags]::None
            )
        }
        
    }

    End {
        $newSddl = $sd.GetSddlForm([System.Security.AccessControl.AccessControlSections]::All)
        if ($newSddl -ne $sddl -and $PSCmdlet.ShouldProcess($Name, "Add DACL entry")) {

            Set-Item -LiteralPath WSMan:\localhost\Service\RootSDDL -Value $newSddl -Force | Out-Null
        }

    }
}

function Find-Group {
    param (
        [string]$GroupName,
        [string]$UserName,
        [string]$DomainName
    )
    $ret = @{
        RC           = $true
        Msg          = ""
        GroupIsLocal = $false
        GroupSID     = $null  
    }
    # Checking the group, we assume is a domain one only if the name is provided as 'Domain\Groupname', if the group is in the domain we assume we have a domain user
    [System.Collections.ArrayList]$MemberList = @()
    if ([bool]$DomainName) {
        try {
            $ADSearcher = New-Object DirectoryServices.DirectorySearcher("(&(objectCategory=group)(sAMAccountName=$GroupName))")
            $SavedEA = $ErrorActionPreference
            $ErrorActionPreference = 'Stop'
            $Results = try {
                $ADSearcher.FindOne()
            }
            catch {
                Write-Host "ERROR querying AD: $_"
            }
            $ErrorActionPreference = $SavedEA

            if (![string]::IsNullOrEmpty($Results)) {
                # Get the group SID
                $groupSID = New-Object System.Security.Principal.SecurityIdentifier($Results.Properties["objectSid"][0], 0)
                $ret.GroupSID = $groupSID.Value  # Store the Group SID
                if ($UserName) {
                    foreach ($r in $Results.Properties["Member"] ) {
                        $MemberList += $(([ADSI]"LDAP://$r").sAMAccountName )
                    }
                    if ($MemberList -contains $UserName) {
                        $ret.Msg += "User $Username is member of group $Groupname. Group SID: $($ret.GroupSID)"
                    }
                    else {
                        $ret.Msg += "User $Username is not member of group $Groupname. Group SID: $($ret.GroupSID) `n Aborting..."
                        $ret.RC = $false
                    }
                }
            }
            else {
                $ret.Msg += "Group $Groupname not found in AD `n Aborting..."
                $ret.RC = $false
            }
        }
        catch {
            Write-Host "ERROR searching AD, make sure the logged in user is a domain member" -ForegroundColor Red -BackgroundColor Black
            $ret.RC = $false
        }
    }
    else {
        $ret.Msg += "Group $Groupname is local"
        $ret.GroupIsLocal = $true
    }
    
    return $ret
}

function Find-User {

    param (
        [string]$UserName,
        [string]$DomainName
    )
    
    $ret = @{
        RC          = $true
        Msg         = ""
        UserIsLocal = $false
    }



    # Checking the user, we aasume is a domain one only if the name is provided as 'Domain\Username'
    # Setting up the WinRM configuration


    if ([bool]$DomainName) {
        $ADSearcher = New-Object DirectoryServices.DirectorySearcher("(&(ObjectClass=User)(sAMAccountName=$UserName))")
   
        if ([bool]$ADSearcher.FindOne()) {
            $ret.Msg += "User $Username found in AD`n" 
        
        }
        else {
            $ret.Msg += "User $Username does not exist, please create it on domain $DomainName`n" 
            $ret.Msg += "Aborting...`n" 
            $ret.RC = $false
        }
       
    }
    else {
        $ret.Msg += "User $UserName is local"
        $ret.UserIsLocal = $true
    }
    return $ret
}

function __ManageLocalUserAndGroup {
    param (
  
        [string]$Username,
        [string]$GroupName,
        [string]$Password,
        [ValidateSet("User", "Group", "All")]
        [string]$Op
    )

    $cmd = "$ENV:Windir\System32\net.exe"
    if (($op -eq "All") -or ($op -eq "User")) {
        # Check if the user exists
        $process = Start-Process -FilePath $cmd -ArgumentList "user $Username "-PassThru  -NoNewWindow -Wait -RedirectStandardError NUL
        $userNotFound = $process.ExitCode 

        if ($userNotFound) {
            # Create the user if it doesn't exist
            Write-Host  "User $UserName does not exist, trying to create it..."
            $createUserProcess = Start-Process -FilePath $cmd -ArgumentList "user $Username $Password /add /Y "-PassThru  -NoNewWindow -Wait  -RedirectStandardError NUL
            if ($createUserProcess.ExitCode -ne 0) {
                write-host "Start-Process -FilePath $cmd -ArgumentList "user $Username $Password /add /Y "-PassThru  -NoNewWindow -Wait  -RedirectStandardError NUL"
                return 1  # Error creating user
            }
        }
    }

    if (($op -eq "All") -or ($op -eq "Group")) {
        # Check if the group exists
        $process = Start-Process -FilePath $cmd -ArgumentList "localgroup $GroupName"-PassThru  -NoNewWindow -Wait -RedirectStandardError NUL 
        $groupNotFound = $process.ExitCode -ne 0

        if ($groupNotFound) {
            # Create the group if it doesn't exist
            Write-Host  "Group $GroupName does not exist, trying to create it..."
            $createGroupProcess = Start-Process -FilePath $cmd -ArgumentList "localgroup $GroupName /add /Y  "-PassThru  -NoNewWindow -Wait -RedirectStandardError NUL 
            if ($createGroupProcess.ExitCode -ne 0) {
                return 2  # Error creating group
            }
        }
        if ($Username) {
            # Check if the user is a member of the group
            $isMember = (net localgroup $GroupName | Select-String -Pattern "^\s*\b$Username\b\s*$" -Quiet)

            if (-not $isMember) {
                # Add the user to the group if they are not a member
                Write-Host  "Adding user $UserName to group $GroupName"
                $addUserToGroupProcess = Start-Process -FilePath $cmd -ArgumentList "localgroup $GroupName $Username /add /Y  "-PassThru  -NoNewWindow -Wait -RedirectStandardError NUL 
                if ($addUserToGroupProcess.ExitCode -ne 0) {
                    Write-Host "Error adding user to group"
                    return 3  # Error adding user to group
                }
            }
        }
    }
    return 0  # Success
} 
function Set-WinRmConfig {
    $RC = [pscustomobject] @{
        output = ''
        result = $false
        
    }
    $local:ErrorActionPreference = 'Stop'
    Write-Host "-> Setting up WINRM service"-ForegroundColor Green -BackgroundColor Black
    Write-Host "Setting network profile"
    try {
        $Profiles = Get-NetConnectionProfile
        foreach ($p in $Profiles) {
            if ($($p.NetworkCategory) -eq "Public") {
                Write-Host "Setting network profile for interface $($p.InterfaceAlias)"
                Set-NetConnectionProfile -NetworkCategory $NetworkProfile 
            }
        }
    }
    catch {
        Write-Host "Error setting network profile"
        Write-Host "Details: $_"
        return $RC.result
    }
            
    try {
        if ((Get-Service WinRM).Status -ne "Running") {

            Enable-PSRemoting -Force
        }
        
        Write-Host "Setting AllowUnencrypted to true"
        winrm set winrm/config/service '@{AllowUnencrypted="true"}' | Out-Null
        [xml]$WinRmConfig = winrm get winrm/config/service -format:pretty
    
        if (([string]::IsNullOrEmpty($WinRmConfig.Service.AllowUnencrypted.Source))) {
            $AllowUnencrypted = ($WinRmConfig.Service.AllowUnencrypted)
        }
        else {
            $AllowUnencrypted = ($WinRmConfig.Service.AllowUnencrypted.'#text')
        }
     
        if (($AllowUnencrypted -eq "true")) {
            $RC.result = $true
        }
        $RC.output = winrm get winrm/config/service
        return $RC
        
    }
    catch {
        Write-Host "Error setting WinRm"
        Write-Host "Details: $_"
        return $RC.result
    }

    
}

function __getpassword {
    Add-Type -AssemblyName System.Web
    $P = ([System.Web.Security.Membership]::GeneratePassword(24, 3))
    $securePassword = ConvertTo-SecureString -String $P -AsPlainText -Force
    return $securePassword
}

function Restart-WinRM {
    Write-Host "-> Restarting WINRM service"
    Stop-Service WinRM | Out-Null
    $MaxWait = 30
    $WmiRestartMsg = $null
    $k = 0
    Write-Host "--> Stopping service"
    do {
        Start-Sleep -Seconds 1
        Write-Host '.' -NoNewline
        $k += 1
        $svcStatus = (Get-Service winrm).status
    } until (
    ($svcStatus -eq "Stopped") -xor ($k -ge $MaxWait)
    )
    Write-Host ''
    if ((Get-Service winrm).status -ne "Stopped") {
        $WmiRestartMsg = "Error stopping winrm service, please investigate"
    }
    else {
        try {
            Write-Host "--> Starting service"
            Start-Service winrm -ea Stop
        
            $startTime = Get-Date
            $timeout = 60 # 60 seconds timeout
            $serviceRunning = $false
        
            while ((Get-Date) -lt $startTime.AddSeconds($timeout)) {
                $service = Get-Service winrm
                if ($service.Status -eq 'Running') {
                    Write-Host "Service winrm is running."
                    $serviceRunning = $true
                    break
                }
                else {
                    Start-Sleep -Seconds 1
                }
            }
        
            if (-not $serviceRunning) {
                throw "Service winrm did not start within the expected time."
            }
        }
        catch {
            $WmiRestartMsg = "Error starting winrm, please investigate and restart"
        }

    }
    
}


#-------------------------------------------------------------------------------
$dscriptver = "1.1"
$VersionBanner = "$($MyInvocation.MyCommand.Name) Version $dscriptver"
$line = (0..$($VersionBanner.length / 2 )) | ForEach-Object { $line + "-" }
Write-Host $line -ForegroundColor White -BackgroundColor Blue
Write-Host "$VersionBanner " -ForegroundColor White -BackgroundColor Blue 
Write-Host $line -ForegroundColor White -BackgroundColor Blue
Write-Host ''
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
try {
    Stop-Transcript -ErrorAction Stop | Out-Null
}
catch {}

$LogFile = "$LogFilePath\$ENV:COMPUTERNAME-$($($MyInvocation.MyCommand.Name).replace('.ps1',''))-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$RC = 1
Start-Transcript -Path $LogFile
if ($UserName -eq $GroupName) {
    Write-Host "Username cannot be the same as GroupName, aborting..."
    return
}

Write-Output "Starting at $(Get-Date)"
Write-Output "Log file is $Logfile"
Write-Output "Windows version: $( (Get-WmiObject -Class Win32_OperatingSystem).Version)"
Write-Output "PS version: $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
Write-Output "`n`n`nParameters as provided --------------------"
foreach ($paramName in $PSBoundParameters.Keys) {
    $paramValue = $PSBoundParameters[$paramName]
    if ($paramName -ne "Pass") {
        Write-Host "Parameter Name: $paramName, Parameter Value: $paramValue "
        Write-Host "-"
    }
}
Write-Output "Environment --------------------"
Write-Output $(Get-ChildItem ENV: | Where-Object { ($_.Name) -NotMatch "DOMOTZ_USER_PASS" })
Write-Output "Log file is $Logfile"
Write-LogLocation $Logfile

Write-Output "This utility will enable WINRM and/or grant WMI permissions on Windows to unlock the Domotz OS Monitoring feature.  (ver. $dscriptver)"
Write-Output "Resolving User and Group"

# Sanitize username and groupname if we are on a domain controller
if ([bool](Get-SmbShare -Name SYSVOL -ea SilentlyContinue)) {
    Write-Host "The computer is a domain controller, assuming group and user are in AD..."
    $NetBIOSName = (Get-ADDomain).NetBIOSName
    if (!($UserName.Contains("\"))) {
        $UserName = "$NetBIOSName\$UserName"
    }
    if (!($GroupName.Contains("\"))) {
        $GroupName = "$NetBIOSName\$GroupName"
    }

}

if ($UserName.Contains("\")) { 
    $UserDomain , $UserSamActName = $UserName.Split("\")
}
else {
    $UserDomain = $null
    $UserSamActName = $UserName

}

if ($GroupName.Contains("\")) { 
    $GroupDomain , $GroupSamActName = $GroupName.Split("\")
}
else {
    $GroupDomain = $null
    $GroupSamActName = $GroupName

}
$ComputerDomain = try {
        (Get-WmiObject Win32_NTDomain).DomainName
}
catch {
    Write-Host "ERROR getting computer domain: $_"
    return $false
}
Write-Host "User Domain is: $UserDomain"
Write-Host "Computer Domain is $ComputerDomain"
if ([bool]$UserDomain) {
    if ((Get-WmiObject Win32_NTDomain).DomainName -notcontains $UserDomain) {
        Write-Warning "The computer is not joint to domain $UserDomain, user and group must be local or belong to the same domain `nAborting..."
        return $false
    }
}

if ([bool]$GroupDomain) {
    if ((Get-WmiObject Win32_NTDomain).DomainName -notcontains $GroupDomain) {
        Write-Warning "The computer is not joint to domain $GroupDomain specified in the GroupName parameter `nAborting..."
        return $false
    }
    Write-Warning "$GroupName is a domain group, assuming the user is in the same domain, password parameter is ignored"
       
    $UserDomain = $GroupDomain
    $UserName = "$UserDomain\$UserSamActName"

}

if (!($PSBoundParameters.ContainsKey('WmiAccessOnly'))) {
    Write-Host "processing user $Username and group $GroupName" -ForegroundColor Green -BackgroundColor Black
    $Fuser = (Find-User -UserName $UserSamActName -DomainName $UserDomain)
    Write-Host $Fuser.Msg
    
    if (($RC = $Fuser.RC)) {
        $FGroup = Find-Group -UserName $UserSamActName -GroupName $GroupSamActName -DomainName $GroupDomain
        Write-Host $FGroup.Msg
        if (($RC = $Fgroup.RC)) {
            # if we have a localgroup we try to add it to the group
            if ($FGroup.GroupIsLocal ) {
            
                if (!$Fuser.UserIsLocal) {

                    # Skip user checking and local creation
                    $op = "Group"
                }
                else {
                    $op = "All"
                }
                if ([string]::IsNullOrEmpty($Pass)) { 
                    [System.Security.SecureString]$securePass = __getpassword

                    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
                    try {
                        $NewPass = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
                    
                    } 
                    catch {
                        Write-Host "Error occurred: $_" -ForegroundColor Cyan -BackgroundColor Black
                    }
                    finally {
                        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

                    }
                }


                $NetCmdRetCode = __ManageLocalUserAndGroup -Username $UserName -GroupName $GroupName -Password $NewPass -Op $op
                switch ($NetCmdRetCode) {
                    0 { "Local user and group operations completed successfully"; break }
                    1 { "Error creating user $UserName" ; break }
                    2 { "Error creating group $GroupName"; break }
                    3 { "Error addin user $UserName to group $GroupName"; break }
                    Default { "Unknown error" }
                }

                try {
                    Start-Sleep -Seconds 5
                    $localGroup = [ADSI]("WinNT://$env:COMPUTERNAME/$GroupName,group")
                    $localGroupSID = New-Object System.Security.Principal.SecurityIdentifier($localGroup.objectsid[0], 0)
                    Write-Host "Local Group SID: $($localGroupSID.Value)"
                }

                catch {
                    Write-Host "Error retrieving Local Group SID: $_"
                }

                if ($NetCmdRetCode -ne 0) {
                    return $false
                }
            }
            else {
                Write-Host "AD Group SID $($FGroup.GroupSID)"
            }

            if ([string]::IsNullOrEmpty($Pass) -and $Fuser.UserIsLocal) {
                Stop-Transcript | Out-Null
                Write-Warning "#################### THIS IS THE GENERATED PASSWORD FOR THE NEW USER, PLEASE TAKE NOTE SINCE IT'S NOT SAVED ANYWHERE`n`n"
                $NewPass
                Write-Host "`n`n"
                Start-Transcript -Path $LogFile -Append | Out-Null

            }

        }
    }

    Write-Host "Configuring WinRM"
    $WinRMConfig = Set-WinRmConfig
    $WinRMConfig.output

    if ($($WinRMConfig.result)) {
        Write-Host "-> Granting WinRM permissions to group $GroupName"
        Add-WinRMDaclRule -Account $GroupName -WhatIf:([bool]$WhatIfPreference.IsPresent) -Confirm:([bool]$ConfirmPreference.IsPresent)
        Restart-WinRM
    }

    else {
        Write-Host "couldn't configure WMI, aborting"
        $RC = $false
    }
}
if ($RC) {
    
    foreach ($n in $Namespaces) {
        Write-Host "-> Granting WMI permissions to $GroupName on namespace $n"
        $RC = $(Set-WMIAcl -Account $GroupName -Namespace $n) -and $RC
    }
}

if ($RC) {
    Write-Host "`n########## The script completed successfully ##########"

    if ($WmiRestartMsg) {
        Write-Host $WmiRestartMsg 
    }
    if (!($PSBoundParameters.ContainsKey('WmiAccessOnly'))) {
        Write-Warning "We have configured WinRM to allow unencrypted authentication, if you want to rollback run the following command:"
        Write-Host "winrm set winrm/config/service '@{AllowUnencrypted=""false""}'"
        Write-Host "Run 'winrm configsddl default' to verify the group has the required permissions"
    }
    else {
        Write-Host "Only WMI permissions have changed"
    }
}
else {
    Write-Host "`nThe script terminated with errors, review the logfile for details"
}


Stop-Transcript -ErrorAction SilentlyContinue 
return [int](!$RC) #I want to return 0 if the ret code is $true