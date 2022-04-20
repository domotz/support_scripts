## Domotz ssh enable PowerShell script for to unock the OS Monitoring feature on your Windows endopoints 


TROUBLESHOOTING
Issues that you might encounter while using these scripts:

1) You get this error:

"the script cannot be loaded because running scripts is disabled on this system...."

Solution: Please run this command (you need administrative privileges)
Set-ExecutionPolicy Unrestricted -Scope LocalMachine

To set it back:
Set-ExecutionPolicy Undefined -Scope LocalMachine


2) You are unable to unlock your device in Domotz even if you have run the script with no errors.

Reason: sometimes user priveleges are not reloaded after the modification.
Solution: Please reboot the Windows Pc to relead the user priveleges and permissions.


3) You are unable to see the OS_Monitoring on SSH option in the Access Manager section

Reason: the agent has not yet discovered the ssh service open on your windows endopoint.
Solution: wait from 1 to 3 hours and check again



It you encounter another isse not listed above please contact support@domotz.com and report it.