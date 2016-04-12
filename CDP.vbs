' ===========================================================================================
'
'   Script Information
'
'   Filename:           getCDPinformations.vbs
'   Author:             Josh Burkard
'   Date:               25.05.2011
'   Description:        get CDP informations from all enabled network adapters
'                       - writes the information to SCCM MIF file
'
' ===========================================================================================
' Constants
' ===========================================================================================

Const NETWORK_CONNECTIONS = &H31&
Const ForAppending        = 8

strScriptPath             = ".\"         ' Path where TCPdump.exe is
strMifFile                = "\CDP.mif"
strComputer               = "."  

' ===========================================================================================
' Check Script is being run with CSCRIPT rather than WSCRIPT due to using stdout
' ===========================================================================================

If UCase(Right(Wscript.FullName, 11)) = "WSCRIPT.EXE" Then
    strPath = Wscript.ScriptFullName
    ' Wscript.Echo "This script must be run under CScript." & vbCrLf & vbCrLf & "Re-Starting Under CScript"
    strCommand = "%comspec% /K cscript //NOLOGO " & Chr(34) & strPath & chr(34)
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run(strCommand)
    Wscript.Quit
End If  

Set objShell = CreateObject("Wscript.Shell")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShellApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' ===========================================================================================
' Get Temporary and MIF-file directory
' ===========================================================================================

strTempDir = objShell.ExpandEnvironmentStrings("%TEMP%")
strWinDir = objShell.ExpandEnvironmentStrings("%WINDIR%")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

' ===========================================================================================
' Get Datas and write it to the MIF-file
' ===========================================================================================

Set objFolder = objShellApp.Namespace(NETWORK_CONNECTIONS)
Set oAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter where NetEnabled = 'True'")
'Set oAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration")

i = 0	' Counter of detected CDP-Informations

' Enumerate the results (list of NICS). "
	Wscript.echo "Scan In Progress" 
For Each oAdapter In oAdapters

    Set oAdapterStingsArray = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where Index= "&oAdapter.DeviceID )
    for each odapterStings in oAdapterStingsArray
				SettingID = odapterStings.SettingID
				Name = oAdapter.NetConnectionID
	Next 
	strCommand = strScriptPath & "\tcpdump -nn -v -s 1500 -i \Device\" & SettingID & " -c 1 ether[20:2] == 0x2000"    
	'Wscript.echo "Executing TCP Dump for Adapter " & chr(34) & oAdapter.SettingID & chr(34) & VbCrLF
	'Wscript.echo strCommand & VbCrLF

	Set objExec = objShell.Exec(strCommand)

	count = 0

	Do Until objExec.Status
		count = count +1
		'Timeout to Deal with Non CDP Enabled Devices
		If count = 250 then
			objExec.terminate
			objExec=""
			wscript.quit
			
		End If
		Wscript.Sleep 250
	Loop    

	' Loop through the output of TCPDUMP stored in stdout and retrieve required fields
	' Namely switch name, IP and Port

    strPort = ""
    strDeviceIP = ""

	While Not objExec.StdOut.AtEndOfStream
		strLine = objExec.StdOut.ReadLine
		If Instr(UCASE(strLine),"DEVICE-ID") > 0 Then
			strDeviceID = Mid(strLine,(Instr(strLine,chr(39))+1),(Len(StrLine) - (Instr(strLine,chr(39))+1)))
		End If
		If Instr(UCASE(strLine),"ADDRESS ") > 0 Then
			strDeviceIP = Right(strLine,(Len(strLine) - (Instrrev(strLine,")")+1)))
		End If
		If Instr(UCASE(strLine),"PORT-ID") > 0 Then
			strPort = Mid(strLine,(Instr(strLine,chr(39))+1),(Len(StrLine) - (Instr(strLine,chr(39))+1)))
		End If
		If Instr(UCASE(strLine),"VLAN ID") > 0 Then
			strVlanArray = Split(strLine,":")
			strVlan = Replace(strVlanArray(2)," ","")
		End If
	Wend   

	If strPort <> "" AND strDeviceIP <> "" Then
		i = i + 1	' Counter of detected CDP-Informations
		Wscript.echo "#######################################################################" 
        wscript.echo "            Value = " & Chr(34) & SettingID & Chr(34)    
        wscript.echo "            Value = " & Chr(34) & Name & Chr(34)
		wscript.echo "            Value = " & Chr(34) & strDeviceID & Chr(34)
		wscript.echo "            Value = " & Chr(34) & strDeviceIP & Chr(34)
		wscript.echo "            Value = " & Chr(34) & strPort & Chr(34)
		wscript.echo "            Value = " & Chr(34) & strVlan & Chr(34)
        strPort = ""
        strDeviceIP = ""
        else 
        Wscript.echo "#######################################################################" 
        wscript.echo "            No Value for "  & Name 
        
	End If

Next 'oAdapter  
