strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")
For Each objPrinter in colInstalledPrinters
  msgbox(objPrinter.Name & " >> " & objPrinter.ShareName& " >> " & objPrinter.DeviceID & " >> " & objPrinter.Local)
  if objPrinter.Name="\\mustafaknt\HPUniver" then
     objPrinter.SetDefaultPrinter()
  end if
Next