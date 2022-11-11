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
.NOTES
        If using a local user that does not exist, the options to assign it a password on creation are the environment variable DOMOTZ_USER_PASS
        or the commandline parameter 'Pass'. If none of those are present the password will be random and revealed on the console.
        The best option is to create the user before running the script so the password will not travel and will be known.
#>


[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory)]
    [string]$UserName,

    [Parameter()]
    [string]
    $Pass = $ENV:DOMOTZ_USER_PASS,

    [Parameter()]
    [string]$GroupName = "DomotzWinRM",

    [Parameter()]
    [ValidateSet("Domain", "Private")]
    [string]$NetworkProfile = "Private"
)

function Set-WMIAcl {
    # Excerpt from https://github.com/grbray/PowerShell/blob/main/Windows/Set-WMINameSpaceSecurity.ps1
    
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Account
    )

    $CONTAINER_INHERIT_ACE_FLAG = 0x2
    $ACCESS_ALLOWED_ACE_TYPE = 0x0
    $WBEM_METHOD_EXECUTE = 0x02
    $WBEM_REMOTE_ACCESS = 0x20


    $ErrorActionPreference = "Stop"

    $InvokeParams = @{Namespace = 'root/cimv2'; Path = "__systemsecurity=@"; ComputerName = $ENV:ComputerName }
    $output = Invoke-WmiMethod @InvokeParams -Name "GetSecurityDescriptor"
    if ($output.ReturnValue -ne 0) { throw "GetSecurityDescriptor failed:  $($output.ReturnValue)" }

    $ACL = $output.Descriptor

    if ($Account.Contains('\')) {
        $Domain = $Account.Split('\')[0]
        if (($Domain -eq ".") -or ($Domain -eq "BUILTIN")) { $Domain = $ComputerName }
        $AccountName = $Account.Split('\')[1]
    }
    elseif ($Account.Contains('@')) {
        $Somain = $Account.Split('@')[1].Split('.')[0]
        $AccountName = $Account.Split('@')[0]
    }
    else {
        $Domain = $ENV:ComputerName
        $AccountName = $Account
    }

    $GetParams = @{Class = "Win32_Account" ; Filter = "Domain='$Domain' and Name='$AccountName'" }
    $Win32Account = Get-WmiObject @GetParams
    if ($Win32Account -eq $null) { throw "Account was not found: $Account" }
   
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

    $output = Invoke-WmiMethod @SetParams
    if ($output.ReturnValue -ne 0) { throw "SetSecurityDescriptor failed: $($output.ReturnValue)" }

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

            Set-Item -LiteralPath WSMan:\localhost\Service\RootSDDL -Value $newSddl -Force
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
        RC  = $true
        Msg = ""
    }
    # Checking the group, we assume is a domain one only if the name is provided as 'Domain\Groupname', if the group is in the domain we assume we have a domain user
    [System.Collections.ArrayList]$MemberList = @()
    if ([bool]$DomainName) {
        # check if the domain is valid
        $ADSearcher = New-Object DirectoryServices.DirectorySearcher("(&(objectCategory=group)(sAMAccountName=$GroupName))")
        $Results = $ADSearcher.FindOne()
        if (![string]::IsNullOrEmpty($Results)) {
            foreach ($r in $Results.Properties["Member"] ) {
                $MemberList += $(([ADSI]"LDAP://$r").sAMAccountName )
            }
            if ($MemberList -contains $UserName) {

                $ret.Msg += "User $Username is member of group $Groupname"
            }
            else {
                $ret.Msg += "User $Username is not member of group $Groupname `n Aborting..."
                $ret.RC = $false
            }
        }
        else {
            $ret.Msg += "Group $Groupname not found in AD `n Aborting..."
            $ret.RC = $false

        }
   
 
    }
    else {
        $ret.Msg += "Group $Groupname is local"
    
        try { 
            Get-LocalGroup $GroupName -Erroraction Stop 
        }
        catch { 
            $ret.Msg += "Group $GrouName not found, creating it locally"
            New-LocalGroup -Name $GroupName -Description "Group for Domotz user" | Out-Null
        }

    }
    
    return $ret
}
function Find-User {

    param (
        [string]$UserName,
        [string]$Pass,
        [string]$DomainName
    )
    
    $ret = @{
        RC        = $true
        Msg       = ""
        RandomPwd = $false
        Password  = ""
    }


    if ([string]::IsNullOrEmpty($Pass)) {
        Add-Type -AssemblyName System.Web
        $Pass = $ret.Password = ([System.Web.Security.Membership]::GeneratePassword(24, 3)) 
        $ret.RandomPwd = $true
    }

    [securestring]$Pass = $($Pass | ConvertTo-SecureString -AsPlainText -Force)
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
        
        try { 
            Get-LocalUser $UserName -Erroraction Stop 
        }
        catch { 
            $ret.Msg += "User $UserName does not exist, let's create it"

            try {
                New-LocalUser $UserName -Password $Pass -Description "Domotz user agent" -ea Stop
            }
            catch {
                $ret.Msg += "$($_.exception.message) Aborting..."
                $ret.RC = $false
            }
        }
    }
    return $ret
}


#-------------------------------------------------------------------------------
$dscriptver = "0.3.6"
$LogFile = "$PSScriptRoot\$($MyInvocation.MyCommand.Name)-$(Get-Date -Format 'yyyy-MM-dd_HH-MM-ss').log"
$RC = 0
Start-Transcript -Path $LogFile
Write-Output "Starting at $(Get-Date)"
Write-Output "Log file is $Logfile"
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
# Check if you have administrative privileges to run this script
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this utility. You should run this with administrative privileges."
    return 1
}
else {
    Write-Information "Code is running as administrator - nice to hear that!"
}

# Sanitize username and groupname if we are on a domain controller
if ([bool](Get-SmbShare -Name SYSVOL -ea SilentlyContinue)) {
    Write-Output "The computer is a domain controller, assuming group and user are in AD..."
    $NetBIOSName = (Get-ADDomain).NetBIOSName
    if (!($UserName.Contains("\"))) {
        $UserName = "$NetBIOSName\$UserName"
    }
    if (!($GroupName.Contains("\"))) {
        $GroupName = "$NetBIOSName\$GroupName"
    }

}
Write-Output "User: $UserName"
Write-Output "Group: $GroupName"

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

if ([bool]$UserDomain) {
    if ((Get-WmiObject Win32_NTDomain).DomainName -ne $UserDomain) {
        Write-Warning "The computer is not joint to domain $UserDomain, user and group must be local or belong to the same domain `nAborting..."
        return 1
    }
}

if ([bool]$GroupDomain) {
    if ((Get-WmiObject Win32_NTDomain).DomainName -ne $GroupDomain) {
        Write-Warning "The computer is not joint to domain $GroupDomain `nAborting..."
        return 1
    }
    Write-Warning "$GroupName is a domain group, assuming the user is in the same domain"
    $UserDomain = $GroupDomain
    $UserName = "$UserDomain\$UserSamActName"
}


$Fuser = (Find-User -UserName $UserSamActName -DomainName $UserDomain -Pass $Pass)
Write-Output $Fuser.Msg

if (($RC = $Fuser.RC)) {
    $FGroup = Find-Group -UserName $UserSamActName -GroupName $GroupSamActName -DomainName $GroupDomain
    Write-Output $FGroup.Msg
    if (($RC = $Fgroup.RC)) {
        # Set network profile
        $Profiles = Get-NetConnectionProfile
        foreach ($p in $Profiles) {
            if ($($p.NetworkCategory) -eq "Public") {
                Set-NetConnectionProfile -NetworkCategory $NetworkProfile
            }
        }

        Write-Output "-> Setting up WINRM service..."
        winrm quickconfig -quiet
        
        # if we have a localgroup we try to add the user to it
        if (![bool]$GroupDomain) {
            if (!(Get-LocalGroupMember -Group $GroupName -member $UserName -EA SilentlyContinue)) {
                try { 
                    Write-Output "Adding $UserName to group $GroupName"
                    Add-LocalGroupMember -Group $GroupName -Member $UserName -Erroraction Stop 
                }
                catch { 
                    Write-Warning "$($_.exception.message) Aborting..."
                    return 1
                }
            }
        }
        Write-Output "-> Restarting WINRM service..."
        Stop-Service WinRM | Out-Null
        $MaxWait = 30
        $WmiRestartMsg = $null
        $k = 0
        Write-Output "--> Stopping service..."
        do {
            Start-Sleep -Seconds 1
            Write-Host '.' -NoNewline
            $k += 1
            $svcStatus = (Get-Service winrm).status
        } until (
            ($svcStatus -eq "Stopped") -xor ($k -ge $MaxWait)
        )
        
        if ((Get-Service winrm).status -ne "Stopped") {
            $WmiRestartMsg = "Error stopping winrm service, please investigate"
        }
        else {
            try {
                Write-Output "--> Starting service..."
                Start-Service winrm -ea Stop
            }
            catch {
                $WmiRestartMsg = "Error starting winrm, please investigate and restart"
            }
        
        }
        
        # Windows needs some time to think about what we just did
        
        (1..30) | ForEach-Object {
            Write-Host '.' -NoNewline
            Start-Sleep 1    
        }

        Write-Output "-> Granting WinRM permissions to group $GroupName"
        Add-WinRMDaclRule -Account $GroupName -WhatIf:([bool]$WhatIfPreference.IsPresent) -Confirm:([bool]$ConfirmPreference.IsPresent)
        Write-Output "-> Granting WMI permissions to $GroupName..."
        Set-WMIAcl -Account $GroupName
        
        if ($Fuser.RandomPwd) {

            Stop-Transcript | Out-Null
            Write-Warning "#################### THIS IS THE GENERATED PASSWORD FOR THE NEW USER, PLEASE TAKE NOTE SINCE IT'S NOT SAVED ANYWHERE`n`n"
            Write-Host $Fuser.Password
            Write-Output "`n`n"
            Start-Transcript -Path $LogFile -Append | Out-Null
            
        }

    }
}

if ($RC) {
    Write-Output "`n########## The script completed successfully ##########"
    Write-Output "Run 'winrm configsddl default' to verify the group has the required permissions"
    if ($WmiRestartMsg) {
        Write-Output $WmiRestartMsg 
    }
}
else {
    Write-Output "`nThe script terminated with errors, review the logfile for details"
}

Stop-Transcript -ErrorAction SilentlyContinue 
return [int]$RC
