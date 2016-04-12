'==========================================================================
'Trend Micro Deep Security Agent installation script 
'Version 1.0
'Date : April 2016
'==========================================================================
on error resume next

'Const
Const ForReading = 1
Const ForWriting = 2 
Const ForAppending = 8
Const ModeAscii = 0
Const HKEY_LOCAL_MACHINE = &H80000002

'Objects
Set WshShell = CreateObject("WScript.Shell")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
Set ObjFSO = CreateObject("Scripting.FileSystemObject")

'Variables
strComputer = wshNetwork.ComputerName
Log_Name = strComputer & "_Install_DeepSecurityInstall.log"
Windir = WshShell.ExpandEnvironmentStrings("%WinDir%")
MSILog_Name = Windir & "\Logs\" & strComputer & "_Install_DeepSecurityInstallMSI.log"
VersionToCheck = "2003"
TrendServer ="ServerName"
TrendServerPort ="4120"
MSI32bits = "Agent-Core-Windows-9.6.2-5451.i386.msi"
MSI64bits = "Agent-Core-Windows-9.6.2-5451.x86_64.msi"

'Logs File Folder Path
If (ObjFSO.FileExists(Windir & "\Logs\" & Log_Name)) Then
	Set MyLogFile = objFSO.OpenTextFile(Windir & "\Logs\" & Log_Name,ForWriting, ModeAscii)
Else
	Set MyLogFile = objFSO.CreateTextFile(Windir & "\Logs\" & Log_Name,ForWriting, ModeAscii)
End If

'Booleans
STATUS_OS_VERSION = False
STATUS_TYPE_DETECT = False
STATUS_INSTALL_AGENT = False
STATUS_FLOWS = False

'#########################################################################
'MAIN : CALL Sub
'-------------------------------------------------------------------------
MyLogFile.writeLine "=================================START================================="
CALL FORCE_CSCRIPT
CALL OSVERSION(VersionToCheck)
MyLogFile.writeLine "=================================END==================================="
MyLogFile.Close

If (STATUS_OS_VERSION = True) and (STATUS_TYPE_DETECT = True) and (STATUS_INSTALL_AGENT = True) and (STATUS_FLOWS = True) Then
	wscript.quit 0
End If

If (STATUS_OS_VERSION = False) Then
	wscript.quit 111
ElseIf (STATUS_TYPE_DETECT = False) Then
	wscript.quit 222
ElseIf (STATUS_INSTALL_AGENT = False) Then
	wscript.quit 333
ElseIf (STATUS_FLOWS = False) Then
	wscript.quit 444
Else
	wscript.quit 555	
End If
'#########################################################################

'==========================================================================
' Force Cscript
'--------------------------------------------------------------------------
Sub FORCE_CSCRIPT
    If right(lCase(wscript.fullname),11)= "wscript.exe" then
        for i=0 to wscript.arguments.count-1
            args = args & wscript.arguments(i) & " "
        next
        set wshshell=createobject("wscript.shell")
        wshshell.run wshshell.ExpandEnvironmentStrings("%comspec%") & " /c cscript.exe //nologo """ & wscript.scriptfullname & """" & args
        set wshshell=nothing
        wscript.quit
    end if
End sub
'=========================================================================

'==========================================================================
'DETECTION OS VERSION
'--------------------------------------------------------------------------
Sub OSVERSION(VersionToCheck)
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 

	If err.number <> 0 then
			MyLogFile.writeLine now & " --> Connexion impossible à WMI : " & err.number & " " & err.description
	End if

	err.clear

	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 

	For Each objItem in colItems 
		If InStr(objItem.Caption,VersionToCheck) Then  
			MyLogFile.writeLine now & " --> OS " & objItem.Caption & " detected"
			STATUS_OS_VERSION = True
			CALL OSTYPE()
		Else
			MyLogFile.writeLine now & " --> OS " & VersionToCheck & " not detected or not supported"
		End If	
	Next
End Sub	
'==========================================================================

'==========================================================================
'DETECTION OS TYPE 32 OR 64 BITS IN REGEDIT
'--------------------------------------------------------------------------
Sub OSTYPE()
	Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			
	If err.number <> 0 Then
		MyLogFile.writeLine now & " --> Registry connection KO ! : " & err.number & " " &  err.description
	Else 
		MyLogFile.writeLine now & " --> Registry connection OK"
	End If

	err.clear
	
	'Search Program Files folder to check OS 32 or 64 bits
	strKeyPathOS = "SOFTWARE\Microsoft\Windows\CurrentVersion"
	strValueNameOS = "ProgramFilesDir (x86)"
			
	objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPathOS,strValueNameOS,strValueOS

	If IsNull(strValueOS) Then
		'OS X86
		MyLogFile.writeLine now & " --> OS X86 detected"
		STATUS_TYPE_DETECT = True
		ExecutablePath = "32bits\" & MSI32bits
		CALL INSTALLAGENT(ExecutablePath)
	Else
		'OS X64
		MyLogFile.writeLine now & " --> OS X64 detected"
		STATUS_TYPE_DETECT = True
		ExecutablePath = "64bits\" & MSI64bits
		CALL INSTALLAGENT(ExecutablePath)
	End If
End Sub	
'==========================================================================

'==========================================================================
'Install Agent
'--------------------------------------------------------------------------
Sub INSTALLAGENT(ExecutablePath)
	
	CurrentFolder = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))	

	If objFSO.FolderExists(CurrentFolder) Then
		'Install Agent
		MSIParameters =  " /quiet /norestart /l*v " & """" & MSILog_Name & """"		
		MyLogFile.writeLine now & " --> Installation in progress : "
		strCommand1 = "msiexec /i " & """" & CurrentFolder & ExecutablePath & """" & MSIParameters 
		MyLogFile.writeLine now & " --> " & strCommand1
		ReturnCode1 = WshShell.Run(strCommand1, 0, True)
		MyLogFile.writeLine now & " --> Return code " & ReturnCode1	
		
		If ReturnCode1 = 0 Then
			STATUS_INSTALL_AGENT = True	
			'Test connectivity
			MyLogFile.writeLine now & " --> Testing Trend Management Server Connectivity : " & TrendServer & ":" & TrendServerPort
			strCommand2 = CurrentFolder & "PortQry.exe"  & " -n " & TrendServer & " -e " & TrendServerPort & " -q"
			MyLogFile.writeLine now & " --> " & strCommand2
			ReturnCode2 = WshShell.Run(strCommand2, 0, True)
			MyLogFile.writeLine now & " --> Return code " & ReturnCode2			
							
			If ReturnCode2 = 0 Then
				'Activate Agent
				MyLogFile.writeLine now & " --> Activation in progress : "
				strCommand3 = """C:\Program Files\Trend Micro\Deep Security Agent\dsa_control.cmd"""  & " -a dsm://" & TrendServer & ":" & TrendServerPort & " --max-dsm-retries 1"
				MyLogFile.writeLine now & " --> " & strCommand3
				'ReturnCode2 = WshShell.Run(strCommand3, 0, True)
				Set oExec = WshShell.Exec (strCommand3)

				Do While Not oExec.StdOut.AtEndOfStream

					strText = oExec.StdOut.ReadLine()
					If Instr(strText, "HTTP Status") > 0 Then
						
						HTTPStatusCode = strText
						MyLogFile.writeLine now & " --> " & HTTPStatusCode
						
						If Instr(HTTPStatusCode, "400") > 0 Then
							MyLogFile.writeLine now & " --> Unable to connect to the Trend Management Server : " & TrendServer & ":" & TrendServerPort & " HTTP Status Code : " & HTTPStatusCode
						ElseIf Instr(HTTPStatusCode, "200") > 0 Then
							MyLogFile.writeLine now & " --> Connect to the Trend Management Server : " & TrendServer & ":" & TrendServerPort & " HTTP Status Code : " & HTTPStatusCode
							STATUS_FLOWS = True
						Else
							MyLogFile.writeLine now & " --> Unable to connect to the Trend Management Server : " & TrendServer & ":" & TrendServerPort & " HTTP Status Code : " & HTTPStatusCode
						End If
						
					Exit Do
					End If
				Loop				
		
			Else
				MyLogFile.writeLine now & " --> Unable to connect to the Trend Management Server : " & TrendServer & ":" & TrendServerPort	
			End If
			
		Else
			MyLogFile.writeLine now & " --> Failed to install the agent."	
		End If
		
	Else
		MyLogFile.writeLine now & " --> Unable to find the folder : " & CurrentFolder		
	End If	

	err.clear
End Sub		
'=========================================================================
