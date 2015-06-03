'Copyright 2012 Google Inc. All Rights Reserved.
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

'GAPS diagnostics tool
'Liron N
Const Ver = "1.0.0"

Dim fso, objShell
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("Wscript.Shell")
Const ForReading = 1, ForWriting = 2, ForAppending = 8 
Const WshFailed = 2
Const WshFinished = 1
Const WshRunning = 0

'Force running in a console window
If Not UCase(Right(WScript.FullName,12))="\CSCRIPT.EXE" Then
	objShell.Run "cscript //nologo """ & WScript.ScriptFullName & """"
	WScript.Quit
End If

On Error Resume Next 'Errors will be handled by code

Dim LogFileName, TempDir
TempDir = objShell.ExpandEnvironmentStrings("%temp%\GAPSTool")
LogFileName = TempDir & "\GAPSTool.log"
'Check if this instance was executed to diagnose a DC
If WScript.Arguments.Count>1 Then
	'Note that this will break commandline arguments if we plan to use them in the future (for UAC for example)
	If UCase(WScript.Arguments(0)) = "/DC" Then
		RunDiagnostics WScript.Arguments(1)
		WScript.Quit
	End If
End If

'Delete old temporary folder
fso.DeleteFolder TempDir,True

'Create new temporary folder
fso.CreateFolder TempDir

LogStr "A:Starting GAPS support tool version " & Ver & " from " & WScript.ScriptFullName

'TODO: Make script request UAC elevation if needed using http://csi-windows.com/toolkit/ifuserperms 
'and http://blogs.technet.com/b/jhoward/archive/2008/11/19/how-to-detect-uac-elevation-from-vbscript.aspx
'for detecting and http://www.winhelponline.com/articles/185/1/VBScripts-and-UAC-elevation.html for requesting.

'Check whether the current user is a Domain Admin, quit if we are not
If Not CheckIfRunningAsDomainAdmin Then WScript.Quit


'Get list of writable DCs
Dim arrWritableDCs
arrWritableDCs = GetWritableDCs
LogStr "A:Got " & UBound(arrWritableDCs)+1 & " writable DCs"
'Instanciate additional arrays
Dim arrExec() 'For Exec objects
ReDim arrExec(UBound(arrWritableDCs))
Dim arrBuffers() 'For StdOut buffers
ReDim arrBuffers(UBound(arrWritableDCs))
Dim arrOutFiles() 'For StdOut buffers
ReDim arrOutFiles(UBound(arrWritableDCs))
For i = 0 To UBound(arrWritableDCs)
	'Create folder for results
	LogStr "A:Creating " & TempDir & "\" & arrWritableDCs(i)
	fso.CreateFolder TempDir & "\" & arrWritableDCs(i)
	If Err<>0 Then LogStr "E:Error creating folder: " & LogErr
	'Call this script with DC name
	LogStr "A:Starting job for " & arrWritableDCs(i)
	'We need to redirect both stdout and stderr to a file instead of catching
	'them directly with the StdOut/StdEr objects, because reading from these
	'streams is blocking, and we want to do it concurrently.
	Set arrExec(i) = objShell.Exec("cmd /c cscript //NoLogo """ & WScript.ScriptFullName & """ /DC " & arrWritableDCs(i) & " 1>" & TempDir & "\" & arrWritableDCs(i) & ".txt 2>&1 ")
	If Err<>0 Then LogStr "E:Error starting job: " & LogErr
	WScript.Sleep 100
	'Open the output file.
	Set arrOutFiles(i) = fso.OpenTextFile(TempDir & "\" & arrWritableDCs(i) & ".txt", ForReading,0)
	If Err<>0 Then LogStr "E:Error opening job output file: " & LogErr
Next 'i

'Process output from all instances until they're all gone
Dim NumCompleted
NumCompleted = 0
While NumCompleted <= UBound(arrWritableDCs)
	WScript.Sleep 10
	For i = 0 To UBound(arrWritableDCs)
		'We set completed Execs to Null, so we can skip them.
		If Not IsNull(arrExec(i)) Then
			arrBuffers(i) = arrBuffers(i) & arrOutFiles(i).Read(1)
			Err.Clear 'Ignore "Input past end of file" errors
      'TODO: Improve logging here - some text files aren't being read on domains with many DCs
			'As long as we have full lines...
			While InStr(arrBuffers(i),vbNewLine)>0 
				If InStr(arrBuffers(i),vbNewLine)>1 Then LogStr "A:Job " & arrWritableDCs(i) & ": " & Left(arrBuffers(i),InStr(arrBuffers(i),vbNewLine)-1)
				arrBuffers(i) = Mid(arrBuffers(i),InStr(arrBuffers(i),vbNewLine)+2,Len(arrBuffers(i)))
			Wend 
			
			If arrExec(i).Status <> WshRunning Then
				'Write any leftover data
				arrBuffers(i) = arrBuffers(i) & arrOutFiles(i).ReadAll
				Err.Clear 'Ignore reading errors
				If Len(arrBuffers(i))>0 Then 
					Dim arrTemp
					arrTemp=Split(arrBuffers(i),vbNewLine)
					For j=0 To UBound(arrTemp)
						If Len(arrTemp(j))>0 Then LogStr "A:Job " & arrWritableDCs(i) & ": " & arrTemp(j)
					Next 'j
				End If
				'Close file we no longer need
				arrOutFiles(i).Close
				'Log status
				If arrExec(i).Status = WshFailed Then
					LogStr "E:Job " & arrWritableDCs(i) & " failed with exit code " & arrExec(i).ExitCode
				ElseIf arrExec(i).Status = WshFinished Then
					LogStr "A:Job " & arrWritableDCs(i) & " finished successfully with exit code " & arrExec(i).ExitCode				
				End If
				NumCompleted=NumCompleted+1
				arrExec(i) = Null
			End If 
		End If	
	Next 'i
Wend
'Report on dll reg/dll loaded/service started issues
'Compare XMLs - they should be identical
'Compare system times if possible
'Ask if there a specific user whose password sync failed
'  Ask when they tried to change their password
'  Get their pwdLastSet, sAMAccountName and "mail" (based on XML)
'    Report when user last changed pwd and if they don't have email address

LogStr "A:Finished collecting information, creating ZIP"

'Create ZIP with reports
Dim ZipName
ZipName = "GAPSTool-report_" & Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & "_" & Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2) & ".zip"
CompressFolder objShell.SpecialFolders("Desktop") & "\" & ZipName,TempDir
WScript.Echo VbNewLine & "Please send the file """ & ZipName & """ from your Desktop to Google Enterprise Support for investigation."
MsgBox "Please send the file """ & ZipName & """ from your Desktop to Google Enterprise Support for investigation.",vbOKOnly,"Google Apps Password Sync diagnostics tool"

WScript.Echo "Press Enter to close this window"
WScript.StdIn.Read(1)

Sub LogStr(str)
	Dim LogFile 'As Stream
	Set LogFile = fso.OpenTextFile(LogFileName,ForAppending,True)
	'TODO: Prettier date/time with DatePart(), and add timezone from http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/daf4b666-fcb6-46ad-becc-689e6daf49ed
	LogFile.WriteLine Now & " " & str
	LogFile.Close
	WScript.Echo Now & " " & str
End Sub

Sub LogErr
	LogErr = "Error #" & Err & " (hex 0x" & Right("00000000" & Hex(Err),8) & "), Source: " & Err.Source & ", Description: " & Err.Description
    Err.Clear
End Sub

'Run diagnostics on remote machines
Sub RunDiagnostics(CompName)
	On Error Resume Next
	Dim LogDir
	LogDir = TempDir & "\" & CompName 
	WScript.StdOut.WriteLine "Starting diagnostics on " & CompName
	objShell.CurrentDirectory = LogDir 'Change current directory
	If Err<>0 Then WScript.StdOut.WriteLine " E:Error changing to work folder for this DC file: " & LogErr
	WScript.StdOut.WriteLine "Getting Notification Package DLL reg entry - dll-reg.txt"
	objShell.Run "cmd /c reg query \\" & CompName & "\HKLM\SYSTEM\CurrentControlSet\Control\Lsa /v ""Notification Packages"" 1>dll-reg.txt 2>dll-reg.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	WScript.StdOut.WriteLine "Getting tasklist.exe output to see if DLL is loaded - dll-loaded.txt"
  'TODO: Verify support for Windows Server 2012 
	objShell.Run "cmd /c tasklist /S " & CompName & " /m password_sync_dll.dll 1>dll-loaded.txt 2>dll-loaded.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	WScript.StdOut.WriteLine "Getting service status - service.txt"
	objShell.Run "cmd /c sc \\" & CompName & " query ""Google Apps Password Sync"" 1>service.txt 2>service.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'  Get logs (from default locations - v1) using XCOPY to get the full tree
	'Assume the username is the same as the current username for the UI logs. Doesn't matter for the others.
	'bWaitOnReturn=True so we don't open too many connections at a time.
	WScript.StdOut.WriteLine "Copying logs and XML - copying.txt"
	'C:\Users\username\AppData\Local\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync"" UI2008 /C /E /F /H /Y /I 1>copying.txt 2>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Documents and Settings\username\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Documents and Settings\%username%\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync"" UI2003 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Users\username\AppData\Local\Google\Identity
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Identity"" Identity2008 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Documents and Settings\username\Local Settings\Application Data\Google\Identity
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Documents and Settings\username\Local Settings\Application Data\Google\Identity"" Identity2003 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service"" Service2008 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service"" Service2003 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass"" DLL2008 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass"" DLL2003 /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\ProgramData\Google\Google Apps Password Sync\config.xml
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\ProgramData\Google\Google Apps Password Sync\config.xml"" /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'C:\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml
	objShell.Run "cmd /c xcopy ""\\" & CompName & "\c$\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml"" /C /E /F /H /Y /I /G 1>>copying.txt 2>>copying.err", 0, True
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'Get install path for GAPS (x86 indicates that the x86 version was installed on x64 - won't work)
	'Just search for the files in both possible paths.
	WScript.StdOut.WriteLine "Getting list of installed files - install.txt and instx86.txt"
	objShell.Run "cmd /c dir ""\\" & CompName & "\c$\Program Files\Google\Google Apps Password Sync"" /B /S 1>install.txt 2>install.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	objShell.Run "cmd /c dir ""\\" & CompName & "\c$\Program Files (x86)\Google\Google Apps Password Sync"" /B /S 1>instx86.txt 2>instx86.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	WScript.StdOut.WriteLine "Geting system-wide proxy settings dump from registry - proxy.txt"
	objShell.Run "cmd /c reg query ""\\" & CompName & "\HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"" /v ProxySettingsPerUser 1>proxy.txt 2>proxy.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	WScript.StdOut.WriteLine "Geting system-wide WinHTTP settings dump from registry - winhttp.txt"
	objShell.Run "cmd /c reg query ""\\" & CompName & "\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"" /v WinHttpSettings 1>winhttp.txt 2>winhttp.err", 0
	If Err<>0 Then WScript.StdOut.WriteLine " E: " & LogErr
	'  Get remote system time using http://blogs.technet.com/b/heyscriptingguy/archive/2007/03/08/how-can-i-verify-the-system-time-on-a-remote-computer.aspx
	WScript.StdOut.WriteLine "Getting local time on remote machine"
	Set objWMIService = GetObject("winmgmts:\\" & CompName & "\root\cimv2")
	If Err<>0 Then WScript.StdOut.WriteLine " E: Error opening WMI on " & CompName & ": " & LogErr

	Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
	If Err<>0 Then WScript.StdOut.WriteLine " E: Error querying Win32_OperatingSystem" & LogErr
	 
	For Each objItem in colItems
		strDate=objItem.LocalDateTime
		WScript.Echo "Local Time: " & Left(strDate, 4) & "-" & Mid(strDate, 7, 2) & "-" & Mid(strDate, 5, 2) & _
			" " & Mid (strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2) & " Time Zone: " & objItem.CurrentTimeZone/60
	Next
	If Err<>0 Then WScript.StdOut.WriteLine " E: Error converting time: " & LogErr
	Err.Clear
	WScript.StdOut.WriteLine "Finished diagnostics on " & CompName
End Sub

'Returns array of writable DCs' DNS names
Function GetWritableDCs()
	On Error Resume Next 
	'Initialize ADSI ADO provider. This is used because we need to make a subtree-scope query.
	Set conn = CreateObject("ADODB.Connection")
	conn.Provider = "ADSDSOObject"
	conn.Open "ADs Provider"
	If Err<>0 Then LogStr "E:Error opening ADSI ADO provider: " & LogErr
	
	'Query for computer accounts where userAccountControl has SERVER_TRUST_ACCOUNT
	'bit set, meaning it's a DC, and not msDS-IsRodc=true, meaning it isn't an RODC.
	'See http://support.microsoft.com/kb/305144 for reference.
	Query = "<LDAP://" & GetObject("LDAP://RootDSE").Get("defaultNamingContext") & ">;" & _
		"(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192)(!(msDS-IsRodc=true)));" & _
		"dNSHostName;subtree"
	
	LogStr "A:Getting list of writable DCs: " & Query
	Set rs = conn.Execute(Query)
	If Err<>0 Then LogStr "E:Error executing query: " & LogErr

	If rs.EOF Then
		LogStr "W:No DCs found - maybe msDS-IsRodc is missing from the schema (Windows 2003)? Trying without it."
		Query = "<LDAP://" & GetObject("LDAP://RootDSE").Get("defaultNamingContext") & ">;" & _
			"(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192));" & _
			"dNSHostName;subtree"
	
		LogStr "A:Getting list of all DCs: " & Query
		Set rs = conn.Execute(Query)
		If Err<>0 Then LogStr "E:Error executing query: " & LogErr

	End If
	
	Dim DCs(), DCCount
	DCCount=0
	While Not rs.EOF
		LogStr "A:Found " & rs.Fields(0).Value
		ReDim Preserve DCs(DCCount)
		DCs(DCCount) = rs.Fields(0).Value
		If Err<>0 Then LogStr "E:Error getting DC name: " & LogErr
		rs.MoveNext
		DCCount=DCCount+1
	Wend
	GetWritableDCs=DCs
End Function

'This function based on http://www.aspfree.com/c/a/Windows-Scripting/Compressed-Folders-in-WSH/
Function CompressFolder(strPath, strFolder)
	On Error Resume Next 
	Const adTypeBinary = 1
	Const adTypeText = 2
	Const adSaveCreateNotExist = 1
	Const adSaveCreateOverwrite = 2
	With CreateObject("ADODB.Stream")
		.Open
		If Err<>0 Then LogStr "E:Error opening ADODB for creating the ZIP file: " & LogErr
		.Type = adTypeText
		.WriteText ChrB(&h50) & ChrB(&h4B) & ChrB(&h5) & ChrB(&h6)
		For i = 1 To 18
			.WriteText ChrB(&h0)
		Next
		.SaveToFile strPath, adSaveCreateNotExist
		If Err<>0 Then LogStr "E:Error saving ZIP file: " & LogErr
		.Close
		.Open
		.Type = adTypeBinary
		.LoadFromFile strPath
		.Position = 2
		arrBytes = .Read
		.Position = 0
		.SetEOS
		.Write arrBytes
		.SaveToFile strPath, adSaveCreateOverwrite
		.Close
		If Err<>0 Then LogStr "E:Error re-saving ZIP file: " & LogErr
	End With
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.NameSpace(strPath)
	If Err<>0 Then LogStr "E:Error opening ZIP file for writing: " & LogErr
	intCount = objFolder.Items.Count
	objFolder.CopyHere strFolder, 256
	If Err<>0 Then LogStr "E:Error copying files to ZIP file: " & LogErr
	Do Until objFolder.Items.Count = intCount + 1
		WScript.Sleep 200
	Loop
End Function

'All functions below either takren from or based on http://explodingcoder.com/blog/content/how-query-active-directory-security-group-membership
' Example code to scan Active Directory for user group membership
' using VBScript under Windows Script Host and ADSI.
'
' Usage: cscript ADGroupsExample.vbs
'
' Shawn Poulson, 2009.05.18
' explodingcoder.com

'Returns True if the current user is a Domain Admin, otherwise False
Function CheckIfRunningAsDomainAdmin()
	'Written by Liron N based on Shawn Poulson
	'NOTE: This function doesn't take into account the actual token's groups,
	'meaning that if running unelevated on a system that uses UAC, the script
	'will not be able to actually use all the user's permissions.
	'TODO: If whoami.exe exists, ignore this completely and only use output from whoami to detect group membership and if UAC elevation is needed.
	On Error Resume Next 
	Set oADSysInfo = CreateObject("ADSystemInfo")
	If Err<>0 Then LogStr "E:Error creating ADSystemInfo object: " & LogErr
	userDN = oADSysInfo.UserName 'Get DN of user
	LogStr "A:Current user DN: " & userDN
	'We shouldn't use the name "Domain Admins" to check membership because it may be localized, we should use the Well-Known SID.
	Const DomainAdminsSIDStart = "010500000000000515000000" 'Domain Admins SID in hex.
	Const DomainAdminsSIDEnd = "00020000" 'Domain Admins SID in hex.
	'Enumerate all member group names
	tkUser = GetTokenGroups(userDN) 'Get tokens of member groups
	'See if the Domain Admins group SID is in the token groups
	CheckIfRunningAsDomainAdmin = False
	For Each sid In tkUser
		Dim tmpstr
		tmpstr=ByteArrToHexString(sid)
		If (Left(tmpstr,Len(DomainAdminsSIDStart)) = DomainAdminsSIDStart) _
			And (Right(tmpstr,Len(DomainAdminsSIDEnd)) = DomainAdminsSIDEnd) Then 
			CheckIfRunningAsDomainAdmin = True
			Exit For
		End If
		If Err<>0 Then LogStr "E:Error checking SID " & tmpstr & ": " & LogErr
	Next 'sid
	If CheckIfRunningAsDomainAdmin Then
		LogStr "A:The current user is a member of Domain Admins"
	Else
		LogStr "E:The current user is *not* a member of Domain Admins"
		MsgBox "Warning: The current user isn't a member of the Domain Admins group. " & _
			"To successfully install and setup Google Apps Directory Sync, you must be a Domain Admin." & _
			vbNewLine & vbNewLine & _
			"Please contact a Domain Admin to continue. You can try running this command, it may add you to the Domain Admins group:" & vbNewLine & vbNewLine & _
			objShell.ExpandEnvironmentStrings("net group ""Domain Admins"" %username% /add") & vbNewLine & vbNewLine & _
			"After joining the Domain Admins group, log out and back in and try again.", _
			vbOKOnly Or vbExclamation, "Google Apps Password Sync diagnostics tool"
		'TODO: Get the correct sAMAccountName for Domain Admins, as it may have been localized...
		'It can be obtained by GetObject("LDAP://<SID=" & ByteArrToHexString(objectSid) & ">").Get("sAMAccountName")	
	End If	
End Function

' Gets tokenGroups attribute from the provided DN
' Usage: <Array of tokens> = GetTokenGroups(<DN of object>)
Function GetTokenGroups(dnObject)
	Dim adsObject
	
	' Setup query of tokenGroup SIDs from dnObject
	Set adsObject = GetObject("LDAP://" & Replace(dnObject, "/", "\/"))
	If Err<>0 Then LogStr "E:Error opening admin's DN using ADSI: " & LogErr
	adsObject.GetInfoEx Array("tokenGroups"), 0
	GetTokenGroups = adsObject.GetEx("tokenGroups")
	If Err<>0 Then LogStr "E:Error getting admin's tokenGroups: " & LogErr
End Function
 
' Encode Byte() to hex string
Function ByteArrToHexString(bytes)
   Dim i
   ByteArrToHexString = ""
   For i = 1 to Lenb(bytes)
      ByteArrToHexString = ByteArrToHexString & Right("0" & Hex(Ascb(Midb(bytes, i, 1))), 2)
      If Err<>0 Then LogStr "E:Error converting SID bytes to string at " & ByteArrToHexString & ": " & LogErr
   Next
End Function
 
'Plans For the future:
'Support non-English systems (i.e. where folder paths are not in English).
'Report which machine the tool is running on
'Gracefully handle case when not running the script as a domain user
'Make an HTML report instead of just collecting text files
'Support paths on upgraded systems such as "C:\WINNT\Profiles\All Users\Application Data\\Google\Google Apps Password Sync\config.xml" etc.
'Make sure all info gathering tasks are done before wrapping up each DC.
'Offer to restart DCs whose DLL is registered but not loaded, if not the current server
'Offer to start the service is it's stopped
'Fix UAC elevation in VBS or remove references to it as it's handled by AutoIt wrapper
'Ask which username is affected and get their LDIF dump, and correlate their pwdLastSet to appearance in the logs - this can tell us if the issue is with the service, the DLL, etc.
'Get the user's LDIF dump using the credentials detailed in the XML
'Try to find password change events in the Event Log for that user to see where password change occurred
'If anonymous access is enabled in the XML, check if it's allowed in AD
'If any of the log files are missing, collect the ACL of that folder (in case there are no permissions for the service user to create the logs). Offer to fix.
'Ask what user was used to install on the other DCs so that we can get the correct path for UI logs, instead of guessing
'Compare XMLs across all servers
'Decode proxy/WinHTTP settings if possible, for example http://p0w3rsh3ll.wordpress.com/2012/10/07/getsetclear-proxy/
'Verify that the BaseDN is correct for this domain
'Get relevant events from Windows Event Logs using "wevtutil"
'Get minidump files: %temp%\WER* folder on Win2008, C:\WINDOWS\pchealth\ERRORREP\UserDumps on Win2003
'Check certificates using certutil -store \\SERVERNAME\AuthRoot | find "Equifax" (or something similar)
'v3:
'Verify that the user used for querying in the XML has correct permissions (warn on every user/OU for which they can't see the "mail" attribute)
'Compare time across DCs
'Compare time to google.com

