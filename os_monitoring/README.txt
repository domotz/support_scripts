
##
If you get this error:

"the script cannot be loaded because running scripts is disabled on this system...."

Please run this command (you need administrative privileges)
Set-ExecutionPolicy Unrestricted -Scope LocalMachine

To set it back:
Set-ExecutionPolicy Undefined -Scope LocalMachine
