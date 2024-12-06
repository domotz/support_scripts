import winrm

s = winrm.Session('127.0.0.1', auth=('testuser','testuser'))
# r = s.run_cmd('ipconfig', ['/all'])
r = s.run_cmd('powershell.exe "Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Manufacturer,Version,BuildNumber,SerialNumber,OSArchitecture | ConvertTo-Json -Compress"')
print(r.std_out)
print(r.std_err)
