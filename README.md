# BuildR

BuildR is a Powershell based application developed by P. Rijal at SICE that uses a C# assembly from itextsharp (buildreport.dll) to generate a pdf build report 
of any devices in a domain environment.

Things to note:

1. BuildR App should run from a domain controller

1. BuildR.exe and buildreport.dll need to exist on the same folder for the application to run

2. Devices should not have their hostnames more than 15 characters because Windows doesn't permit computer names 
that exceed 15 characters. It will skip those devices and complete the buildreport. To generate BR for those devices, please download BuildR app 
to the skipped computers and use CLIENT tab in BuildR to generate report locally.

3. Tab Windows: This tab is used to build report for all windows devices on a domain environment. It should work on any projects without any change in the script.

3. Linux Tab: This tab is used to build report for esxi hosts or any linux VMs. Depending on the project requirements, the script will need minor ammendments for 100% correct results.

4. Client Tab: explained on point number 2. To use this tab, the application folder must be downloaded on the local disk/desktop.

If you come across any issues, please feel free to reach out to: 

Pawan Rijal
pawanrijal@sice.com.au
0404618182
