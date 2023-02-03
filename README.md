# buildreport
Buildreport for the devices in a domain


Buildreport is a Powershell based application developed at SICE by Pawan Rijal that uses a C# assembly from itextsharp (buildreport.dll) to generate a pdf build report 
of any devices in a domain environment.

Few things to note:

1. Run the buildreport application as an administrator in a Domain Controller, the script uses multiple cmdlets that might need admin rights.

2. Buildreport.exe and buildreport.dll need to exist on the same folder for the application to run

3. Devices should not have their hostnames more than 15 characters because Windows doesn't permit computer names 
that exceed 15 characters. It will skip those devices and complete the buildreport. You could use the
exclusive.ps1 script for those devices and run it locally.

If you come across any issues, please feel free to reach out to: 

Pawan Rijal
pawanrijal@outlook.com
