Domotz Diagnostic Powershell Script


TROUBLESHOOTING
1) if you get this king of error:

.\domotz_win_diagnostics.ps1 : File
C:\...\domotz_win_diagnostics.ps1 cannot be loaded
because running scripts is disabled on this system. For more information, see about_Execution_Policies at
https:/go.microsoft.com/fwlink/?LinkID=135170.
At line:1 char:1
+ .\domotz_win_diagnostics.ps1
+ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : SecurityError: (:) [], PSSecurityException
    + FullyQualifiedErrorId : UnauthorizedAccess


Please follow these instructions:

1) open a PowerShell terminal with Adminstrative privileges and allow the script to be executed by issuing this command:
Set-ExecutionPolicy unrestricted

2) run the script again

3) After it has completed all his steps and created a .zip file on your Desktop, please revert the PS ExecutionPolify back to restricted:

Set-ExecutionPolicy restricted
