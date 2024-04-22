## Domotz WINRM enable PowerShell script to unlock the OS Monitoring feature on your Windows endpoints 

To see the help please do:
Unblock-File '.\enable_winrm_os_monitoring.ps1'
man '.\enable_winrm_os_monitoring.ps1'

To see examples on ps:
man .\enable_winrm_os_monitoring.ps1 -examples

How to use:

-------------------------- EXAMPLE 0 --------------------------
.\enable_winrm_os_monitoring_new.ps1 domotzuser

check if the local domotzuser username exists, if not it creates it
check if the local DomotzWinRM group, if not creates it

-------------------------- EXAMPLE 1 --------------------------

.\enable_winrm_os_monitoring_new.ps1 -UserName domotz\domotztestuser -GroupName domotz\ddomaingrp

Checks if the group exists in AD and the user is a member of the group, if not it terminates (no attempt to create
objects in AD are made by the script).
    If the group exists and the user is a member of the group the script grants permissions to the group on the
WinRM default listener

-------------------------- EXAMPLE 2 --------------------------
.\enable_winrm_os_monitoring_new.ps1 -UserName domotzlocaluser -GroupName domotz\ddomaingrp

Since the group is a domain one, the script assumes the user is in the same domain and a member of the group

-------------------------- EXAMPLE 3 --------------------------

.\enable_winrm_os_monitoring_new.ps1 -UserName domotzlocaluser -GroupName domotzLocalGroup

Group and user will be created locally if missing, the user will be added to the group if not there already and
permissions will be granted to the group


N.B. remember to change the user password if you created a new user from scratch.

-------------------------- EXAMPLE 4 --------------------------

.\enable_winrm_os_monitoring_new.ps1 -UserName domotz\domotztestuser -GroupName ddomaingrp

checks if the uesr exists in AD
checks if the group exists locally since no domain is provided
add the user to the local group and grant the group permissions on the default WinRM listener
 

-------------------------- EXAMPLE 5 --------------------------

.\enable_winrm_os_monitoring.ps1 -GroupName adlab\domotzwinrm -WmiAccessOnly -Namespaces "Root\Microsoft\Windows\Storage"

Add WMI permissions to adlab\domotzwinrm group on namespace "Root\Microsoft\Windows\Storage"


TROUBLESHOOTING
Issues that you might encounter while using these scripts:

1) You get this error:

"the script cannot be loaded because running scripts is disabled on this system...."

Solution: Please run this command (you need administrative privileges)
Set-ExecutionPolicy Unrestricted -Scope LocalMachine

To set it back:
Set-ExecutionPolicy Undefined -Scope LocalMachine


2) You are unable to unlock your device in Domotz even after running the script with no errors.

Reason: sometimes user privileges are not reloaded after the modification.
Solution: Please reboot the Windows Pc to reload the user privileges and permissions.


3) You are unable to get help when performing: Get-Help .\enable_winrm_os_monitoring.ps1

You need to unblock the file: please perform
Unblock-File '.\enable_winrm_os_monitoring.ps1'

Then do Get-Help .\enable_winrm_os_monitoring.ps1 again
