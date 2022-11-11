## Domotz SSH enable PowerShell script to unlock the OS Monitoring feature on your Windows endpoints 


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


3) You are unable to see the OS_Monitoring on SSH option in the Access Manager section

Reason: the agent has not yet discovered the SSH service open on your Windows endpoint.
Solution: wait up to 3 hours and check again. The agent may take up to 3 hours to detect the SSH running on the endpoint.

If you encounter another issue not listed above, please email support@domotz.com and report it.