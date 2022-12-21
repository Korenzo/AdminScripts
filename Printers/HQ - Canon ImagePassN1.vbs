'VBScript that installs ImagePass-N1 Copier in the Philadelphia Office.
' - Will delete the old printer
' - Installs the new printer for Win7 x64
' - Installs the driver
' - 10.1.2.209 - Philadelphia office Color Copier

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    objWMIService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

 ' QUERY THE USER OS        
Set colOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystem
Version = objOperatingSystem.Version

' QUERY TO SEE IF PRINTER EXISTS
Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where PortName = 'IP_10.1.2.209' or PortName = '10.1.2.209'")
' IF SO, DELETE THE PRINTER
For Each objPrinter in colPrinters
    objPrinter.Delete_
Next

' QUERY TO SEE IF PORT EXISTS
Set colInstalledPorts =  objWMIService.ExecQuery _
    ("Select * from Win32_TCPIPPrinterPort Where Name = 'IP_10.1.2.209' or Name ='10.1.2.209'")
' IF SO, DELETE THE PORT
For Each objPort in colInstalledPorts
ON ERROR RESUME NEXT
    objPort.Delete_
Next


IF Mid(Version,1,3) = "6.1" Or Mid(Version,1,4) = "10.0" Then

Bits = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth

	IF Bits = "64" Then

        ' INSTALLS DRIVER
    strINFPath = "\\urbanengineers.local\files\legacy\XFERS\IT\Drivers\Canon\ImagePass-N1\x64\pcl6\Cnp60MA64.INF"
    strPrinterName = "Canon Generic Plus PCL6"
    strArch = "x64"
    strMode = "Type 3 - User Mode"
    Set WshShell = CreateObject("WSCript.shell")
    strCommand = "printui.exe /ia /m """ & strPrinterName & """ /h """ & strArch & """ /v """ & strMode & """ /f " & strINFPath & ""
    Dim result
    result = WshShell.Run(strCommand, 1, True)

	'CREATE PORT
    Set objNewPort = objWMIService.Get _ 
        ("Win32_TCPIPPrinterPort").SpawnInstance_ 
    objNewPort.Name = "10.1.2.209" 
    objNewPort.Protocol = 1 
    objNewPort.HostAddress = "10.1.2.209" 
    objNewPort.PortNumber = "9100"
    objNewPort.SNMPEnabled = False 
    objNewPort.Put_ 
 
    'SETS COPIER TO PORT
    Set objPrinter = objWMIService.Get _ 
        ("Win32_Printer").SpawnInstance_ 
    objPrinter.DriverName = "Canon Generic Plus PCL6" 
    objPrinter.PortName   = "10.1.2.209"
    objPrinter.DeviceID   = "Canon irADVC7565 (Corp Dev)"
    objPrinter.Location = "Philadelphia Office" 
    objPrinter.Network = True 
    objPrinter.Shared = False 
    'objPrinter.ShareName = 
    objPrinter.Put_ 

	ELSE

MsgBox "You have a x86 version of Windows. This script does not support your computer setup."
Wscript.Quit

	End If

ELSE

MsgBox "This Version of Windows is not Supported. Contact your administrator. "
Wscript.Quit

END If

Next

MsgBox "The Canon irADVC7565 - Color is now installed."