' Copyright 2012 Google Inc. All Rights Reserved.
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' GSPS diagnostics tool
' Liron Newman lironn@google.com

' Do not change this line's format, build.bat relies on it.
Const Ver = "2.0.0.0"

Dim fso, objShell, CurrentComputerName
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("Wscript.Shell")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const WshRunning = 0, WshFinished = 1, WshFailed = 2
Const HKEY_LOCAL_MACHINE = &H80000002  ' From https://msdn.microsoft.com/en-us/library/aa394600(v=vs.85).aspx?cs-lang=vb

' Force running in a console window
If Not UCase(Right(WScript.FullName, 12)) = "\CSCRIPT.EXE" Then
  objShell.Run "cscript //nologo """ & WScript.ScriptFullName & """"
  WScript.Quit
End If

On Error Resume Next  ' Errors will be handled by code

Dim LogFileName, TempDir
TempDir = objShell.ExpandEnvironmentStrings("%temp%\GSPSTool")
' We can assume %userdnsdomain% is the computer's DNS domain too, because we
' enforce it in CheckComputerAndUserDetails().
CurrentComputerName = _
    UCase(objShell.ExpandEnvironmentStrings("%computername%.%userdnsdomain%"))
LogFileName = TempDir & "\GSPSTool.log"
' Check if this instance was executed to diagnose a DC
If WScript.Arguments.Count > 1 Then
  ' Note that this will break commandline arguments if we plan to use them in
  ' the future
  If UCase(WScript.Arguments(0)) = "/DC" Then
    RunDiagnostics WScript.Arguments(1)
    WScript.Quit
  End If
End If

' Delete old temporary folder
fso.DeleteFolder TempDir, True

' Create new temporary folder
fso.CreateFolder TempDir

LogStr "A:Starting GSPS support tool version " & Ver & " from " & _
       WScript.ScriptFullName

' Check whether the current user is a Domain Admin and other machine/user
' settings.
If Not CheckComputerAndUserDetails Then WScript.Quit
If Not CheckIfRunningAsDomainAdmin Then WScript.Quit


' Get list of writable DCs
Dim arrWritableDCs
arrWritableDCs = GetWritableDCs
LogStr "A:Got " & UBound(arrWritableDCs) + 1 & " writable DCs"
' Instanciate additional arrays
Dim arrExec()  ' For Exec objects
ReDim arrExec(UBound(arrWritableDCs))
Dim arrBuffers()  ' For StdOut buffers
ReDim arrBuffers(UBound(arrWritableDCs))
Dim arrOutFiles()  ' For StdOut buffers
ReDim arrOutFiles(UBound(arrWritableDCs))
For i = 0 To UBound(arrWritableDCs)
  ' Create folder for results
  LogStr "A:Creating " & TempDir & "\" & arrWritableDCs(i)
  fso.CreateFolder TempDir & "\" & arrWritableDCs(i)
  LogErrorIfNeeded "Error creating folder"
  ' Call this script with DC name
  LogStr "A:Starting job for " & arrWritableDCs(i)
  ' We need to redirect both stdout and stderr to a file instead of catching
  ' them directly with the StdOut/StdEr objects, because reading from these
  ' streams is blocking, and we want to do it concurrently.
  Set arrExec(i) = objShell.Exec("cmd /c cscript //NoLogo """ & _
                                 WScript.ScriptFullName & _
                                 """ /DC " & _
                                 arrWritableDCs(i) & _
                                 " 1>" & _
                                 TempDir & _
                                 "\" & _
                                 arrWritableDCs(i) & _
                                 ".txt 2>&1 ")
  LogErrorIfNeeded "Error starting job"
  WScript.Sleep 100
  ' Open the output file.
  Set arrOutFiles(i) = _
      fso.OpenTextFile(TempDir & "\" & arrWritableDCs(i) & ".txt", _
                       ForReading, _
                       0)
  LogErrorIfNeeded "Error opening job output file"
Next

' Process output from all instances until they're all gone
Dim NumCompleted
NumCompleted = 0
While NumCompleted <= UBound(arrWritableDCs)
  WScript.Sleep 10
  For i = 0 To UBound(arrWritableDCs)
    ' We set completed Execs to Null, so we can skip them.
    If Not IsNull(arrExec(i)) Then
      arrBuffers(i) = arrBuffers(i) & arrOutFiles(i).Read(1)
      Err.Clear  ' Ignore "Input past end of file" errors
      ' TODO: Improve logging here - some text files aren't being read on
      ' domains with many DCs
      ' As long as we have full lines...
      While InStr(arrBuffers(i), vbNewLine) > 0
        If InStr(arrBuffers(i), vbNewLine) > 1 Then
          LogStr "A:Job " & arrWritableDCs(i) & ": " & _
                 Left(arrBuffers(i), InStr(arrBuffers(i), vbNewLine) - 1)
        End If
        arrBuffers(i) = Mid(arrBuffers(i), _
                            InStr(arrBuffers(i), vbNewLine) + 2, _
                            Len(arrBuffers(i)))
      Wend

      If arrExec(i).Status <> WshRunning Then
        ' Write any leftover data
        arrBuffers(i) = arrBuffers(i) & arrOutFiles(i).ReadAll
        Err.Clear  ' Ignore reading errors
        If Len(arrBuffers(i)) > 0 Then
          Dim arrTemp
          arrTemp = Split(arrBuffers(i), vbNewLine)
          For j = 0 To UBound(arrTemp)
            If Len(arrTemp(j)) > 0 Then
              LogStr "A:Job " & arrWritableDCs(i) & ": " & arrTemp(j)
            End If
          Next  ' j
        End If
        ' Close file we no longer need
        arrOutFiles(i).Close
        ' Log status
        If arrExec(i).Status = WshFailed Then
          LogStr "E:Job " & arrWritableDCs(i) & " failed with exit code " & _
                 arrExec(i).ExitCode
        ElseIf arrExec(i).Status = WshFinished Then
          LogStr "A:Job " & arrWritableDCs(i) & _
                 " finished successfully with exit code " & arrExec(i).ExitCode
        End If
        NumCompleted = NumCompleted + 1
        arrExec(i) = Null
      End If
    End If
  Next  ' i
Wend

LogStr "A:Finished collecting information, creating ZIP"

' Create ZIP with reports
Dim ZipName
ZipName = "GSPSTool-report_" & _
          Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
          "_" & _
          Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & _
          Right("0" & Second(Now), 2) & ".zip"
CompressFolder objShell.SpecialFolders("Desktop") & "\" & ZipName, TempDir
Message = "Please send the file """ & ZipName & _
          """ from your Desktop to Google Cloud Support for investigation."
WScript.Echo VbNewLine & Message
MsgBox Message, vbOKOnly, "G Suite Password Sync diagnostics tool"

WScript.Echo "Press Enter to close this window"
WScript.StdIn.Read(1)


Sub LogStr(str)
  Dim LogFile  ' As Stream
  Set LogFile = fso.OpenTextFile(LogFileName, ForAppending, True)
  ' TODO: Prettier date/time with DatePart(), and add timezone from http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/daf4b666-fcb6-46ad-becc-689e6daf49ed
  LogFile.WriteLine Now & " " & str
  LogFile.Close
  WScript.Echo Now & " " & str
End Sub

Sub LogErr
  LogErr = "Error #" & Err & " (hex 0x" & Right("00000000" & Hex(Err), 8) & _
           "), Source: " & Err.Source & ", Description: " & Err.Description
  Err.Clear
End Sub

Sub PrintLine(Text)
  WScript.StdOut.WriteLine Text
End Sub

Sub PrintErrorIfNeeded(Text)
  If Err <> 0 Then PrintLine " E:" & Text & " " & LogErr
End Sub

Sub LogErrorIfNeeded(Text)
  If Err <> 0 Then LogStr "E:" & Text & ": " & LogErr
End Sub

Sub ErrorMsgBox(Text)
  MsgBox "Error: " & Text, _
         vbOKOnly Or vbExclamation, _
         "G Suite Password Sync diagnostics tool"
End Sub

Sub RunCommand(Command, OutputFileNameBase)
  On Error Resume Next

  ' Always use bWaitOnReturn=True to make sure the subpreoccess returns after
  ' all data was collected.
  PrintLine "Running command: " & Command
  objShell.Run "cmd /c " & Command & " 1>>" & OutputFileNameBase & ".txt " & _
                   "2>>" & OutputFileNameBase & ".err", _
               0, _
               True
  PrintErrorIfNeeded "Running command '" & Command & "' failed. "
End Sub

Sub RunCopyCommand(Source, Target)
  ' Checking if we are copying from the local machine and current user.
  ' If we are, use %userprofile% which is more reliable.
  CurrentMachineAndUserPrefix2008 = _
      "\\" & CurrentComputerName & "\C$\USERS\%USERNAME%\"
  CurrentMachineAndUserPrefix2003 = _
      "\\" & CurrentComputerName & "\C$\DOCUMENTS AND SETTINGS\%USERNAME%\"
  If UCase(Left(Source, Len(CurrentMachineAndUserPrefix2008))) = _
      CurrentMachineAndUserPrefix2008 Then
    Source = "%userprofile%" & _
             Mid(Source, Len(CurrentMachineAndUserPrefix2008))
  ElseIf UCase(Left(Source, Len(CurrentMachineAndUserPrefix2003))) = _
      CurrentMachineAndUserPrefix2003 Then
    Source = "%userprofile%" & _
             Mid(Source, Len(CurrentMachineAndUserPrefix2003))
  End If

  RunCommand "xcopy """ & Source & """ """ & Target & """ " & _
                 "/C /E /F /H /Y /I /G", _
             "copying"
End Sub

Sub DecodeWinHTTPSettings(CompName, OutputFileName)
  On Error Resume Next

  LogLinePrefix = "Current WinHTTP proxy settings:" & vbCRLF & vbCRLF
  ' Create a WMI StdRegProv.
  Dim objStdRegProv
  Set objStdRegProv = GetObject( _
      "winmgmts:{impersonationLevel=impersonate}!\\" & _
      CompName & "\root\default:StdRegProv")
  PrintErrorIfNeeded "Error opening WMI StdRegProv on " & CompName & ": "
  ' Retrieve the value of WinHTTPSettings from the registry.
  ' Note that GetBinaryValue returns an array, where each element in the array
  ' is a DECIMAL value of the octets.
  Dim WinHTTPSettingsArray
  objStdRegProv.GetBinaryValue HKEY_LOCAL_MACHINE, _
                               "SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
                                   "Internet Settings\Connections", _
                               "WinHttpSettings", _
                               WinHTTPSettingsArray
  PrintErrorIfNeeded "Error retrieving HKLM\SOFTWARE\Microsoft\Windows\" & _
      "CurrentVersion\Internet Settings\Connections\WinHttpSettings on " & _
      CompName & ": "
  ' The WinHttpSettings registry value appears to be formatted as follows:
  '   Length : Description
  '       12 : ?
  '        1 : Length of proxy string.
  '        3 : ?
  '        ~ : Proxy string; variable length.
  '        1 : Length of bypass list string.
  '        3 : ?
  '        ~ : Bypass list string; variable length.
  ' Based on https://p0w3rsh3ll.wordpress.com/2012/10/07/getsetclear-proxy/.
  ' Start by getting the proxy string length.
  Dim WinHTTPProxyLength
  WinHTTPProxyLength = WinHTTPSettingsArray(12)
  ' Prepare the output file.
  Dim WinHTTPParsedFile
  Set WinHTTPParsedFile = fso.OpenTextFile(OutputFileName, ForAppending, True)
  PrintErrorIfNeeded "Error opening " & OutputFileName
  ' If the proxy string length is greater than 0, a proxy is set. If not, the
  ' connection is direct.
  If WinHTTPProxyLength > 0 Then
    Dim WinHTTPProxy, WinHTTPBypassList, WinHTTPBypassListLength
    ' Concatenate the proxy, starting from 16, through the proxy length.
    For Index = 16 To (16 + WinHTTPProxyLength - 1)
      WinHTTPProxy = WinHTTPProxy & ChrW(WinHTTPSettingsArray(Index))
    Next
    ' Get the bypass list string length. We know its position is 12 + 1 + 3 +
    ' the length of the proxy string. 
    WinHTTPBypassListLength = WinHTTPSettingsArray((16 + WinHTTPProxyLength))
    ' If the length of the list is greater than 0, concatenate it.
    If WinHTTPBypassListLength > 0 Then
      ' Start from 12 + 1 + 3 + proxy string length + 1 + 3.
      For Index = (20 + WinHTTPProxyLength) To _
          (20 + WinHTTPProxyLength + WinHTTPBypassListLength - 1)
        WinHTTPBypassList = WinHTTPBypassList & _
                            ChrW(WinHTTPSettingsArray(Index))
      Next
    Else
      WinHTTPBypassList = "(none)"
    End If
    PrintErrorIfNeeded "Error decoding WinHttpSettings on " & CompName & ": "
    WinHTTPParsedFile.WriteLine LogLinePrefix & _
        "    Proxy Server(s) :  " & WinHTTPProxy & vbCRLF & _
        "    Bypass List     :  " & WinHTTPBypassList
  Else
    WinHTTPParsedFile.WriteLine LogLinePrefix & _
        "    Direct access (no proxy server)."
  End If
  PrintErrorIfNeeded "Error writing to " & OutputFileName
  WinHTTPParsedFile.Close
End Sub

' Run diagnostics on remote machines
Sub RunDiagnostics(CompName)
  On Error Resume Next

  PrintLine "Starting diagnostics on " & CompName
  objShell.CurrentDirectory = TempDir & "\" & CompName  ' Change current dir
  PrintErrorIfNeeded "Error changing to work folder for this DC file: "

  PrintLine "Getting Notification Package DLL reg entry - dll-reg.txt"
  RunCommand "reg query \\" & CompName & _
                 "\HKLM\SYSTEM\CurrentControlSet\Control\Lsa " & _
                 "/v ""Notification Packages""", _
             "dll-reg"

  PrintLine "Running tasklist.exe to see if the DLL is loaded - dll-loaded.txt"
  RunCommand "tasklist /S " & CompName & " /m password_sync_dll.dll", _
             "dll-loaded"

  PrintLine "Getting service status - service_gaps.txt and service_gsps.txt"
  RunCommand "sc \\" & CompName & " query ""Google Apps Password Sync""", _
             "service_gaps"
  RunCommand "sc \\" & CompName & " query ""G Suite Password Sync""", _
             "service_gsps"

  ' Get logs (from default locations - v1) using XCOPY to get the full tree
  ' Assume the username is the same as the current username for the UI logs.
  ' It doesn't matter for the other paths (they don't depend on the username).
  PrintLine "Copying logs and XML - copying.txt"

  ' C:\Users\username\AppData\Local\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync
  RunCopyCommand "\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync", _
                 "UI"

  ' C:\Documents and Settings\username\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\%username%\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\GoogleAppsPasswordSync", _
                 "UI"

  ' C:\Users\username\AppData\Local\Google\Google Apps Password Sync\Tracing\PasswordSync
  RunCopyCommand "\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Google Apps Password Sync\Tracing\PasswordSync", _
                 "UI"

  ' C:\Documents and Settings\username\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\PasswordSync
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\%username%\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\PasswordSync", _
                 "UI"

  ' C:\Users\username\AppData\Local\Google\Identity
  RunCopyCommand "\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Identity", _
                 "Identity"

  ' C:\Documents and Settings\username\Local Settings\Application Data\Google\Identity
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\username\Local Settings\Application Data\Google\Identity", _
                 "Identity"

  ' C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service
  RunCopyCommand "\\" & CompName & "\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service", _
                 "Service"

  'C:\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service", _
                 "Service"

  ' C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Identity
  RunCopyCommand "\\" & CompName & "\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Identity", _
                 "ServiceAuth"

  'C:\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Identity
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Identity", _
                 "ServiceAuth"

  ' C:\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass
  RunCopyCommand "\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass", _
                 "DLL"

  'C:\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass
  RunCopyCommand "\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass", _
                 "DLL"

  ' C:\ProgramData\Google\Google Apps Password Sync\config.xml
  RunCopyCommand "\\" & CompName & "\c$\ProgramData\Google\Google Apps Password Sync\config.xml", _
                 "."

  ' C:\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml
  RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml", _
                 "."

  ' Get install path for GSPS (x86 indicates that the x86 version was installed
  ' on x64 - won't work). Just search for the files in both possible paths.
  PrintLine "Getting list of installed files - install.txt and instx86.txt"
  RunCommand "dir ""\\" & CompName & "\c$\Program Files\Google\Google Apps Password Sync"" /B /S", _
             "install"
  RunCommand "dir ""\\" & CompName & "\c$\Program Files\Google\G Suite Password Sync"" /B /S", _
             "install"
  RunCommand "dir ""\\" & CompName & "\c$\Program Files (x86)\Google\Google Apps Password Sync"" /B /S", _
             "instx86"
  RunCommand "dir ""\\" & CompName & "\c$\Program Files (x86)\Google\G Suite Password Sync"" /B /S", _
             "instx86"

  PrintLine "Getting system-wide proxy settings dump from registry - proxy.txt"
  RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"" /v ProxySettingsPerUser", _
             "proxy"

  PrintLine "Getting system-wide WinHTTP settings dump from registry - winhttp.txt"
  RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"" /v WinHttpSettings", _
             "winhttp"

  PrintLine "Getting admin email address and service account address (if applicable)"
  RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v Email", _
             "admin-and-serviceaccount-emails"
  RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v ServiceAccountEmail", _
             "admin-and-serviceaccount-emails"

  PrintLine "Getting system-wide WinHTTP settings dump from registry, and decoding - winhttp_decoded.txt"
  DecodeWinHTTPSettings CompName, "winhttp_decoded.txt"

  ' Get remote system time using http://blogs.technet.com/b/heyscriptingguy/archive/2007/03/08/how-can-i-verify-the-system-time-on-a-remote-computer.aspx
  PrintLine "Getting local time on remote machine"
  Set objWMIService = GetObject("winmgmts:\\" & CompName & "\root\cimv2")
  PrintErrorIfNeeded "Error opening WMI on " & CompName & ": "

  Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
  PrintErrorIfNeeded "Error querying Win32_OperatingSystem. "

  For Each objItem in colItems
    strDate = objItem.LocalDateTime
    WScript.Echo "Local Time: " & _
                 Left(strDate, 4) & "-" & _
                 Mid(strDate, 7, 2) & "-" & _
                 Mid(strDate, 5, 2) & " " & _
                 Mid (strDate, 9, 2) & ":" & _
                 Mid(strDate, 11, 2) & ":" & _
                 Mid(strDate, 13, 2) & _
                 ", Time Zone: " & (objItem.CurrentTimeZone / 60)
  Next
  PrintErrorIfNeeded "Error converting time: "
  Err.Clear

  PrintLine "Finished diagnostics on " & CompName
End Sub

' Returns array of writable DCs' DNS names
Function GetWritableDCs()
  On Error Resume Next

  ' Initialize ADSI ADO provider. This is used because we need to make a
  ' subtree-scope query.
  Set conn = CreateObject("ADODB.Connection")
  conn.Provider = "ADSDSOObject"
  conn.Open "ADs Provider"
  LogErrorIfNeeded "Error opening ADSI ADO provider"

  QueryBase = "<LDAP://" & _
              GetObject("LDAP://RootDSE").Get("defaultNamingContext") & ">;"
  ' Query for computer accounts where userAccountControl has
  ' SERVER_TRUST_ACCOUNT bit set, meaning it's a DC, and not msDS-IsRodc=true,
  ' meaning it isn't an RODC. See http://support.microsoft.com/kb/305144 for
  ' reference.
  Query = QueryBase & _
          "(&(objectCategory=computer)" & _
          "(userAccountControl:1.2.840.113556.1.4.803:=8192)" & _
          "(!(msDS-IsRodc=true)));" & _
          "dNSHostName;subtree"

  LogStr "A:Getting list of writable DCs: " & Query
  Set rs = conn.Execute(Query)
  LogErrorIfNeeded "Error executing query"

  If rs.EOF Then
    LogStr "W:No DCs found - maybe msDS-IsRodc is missing from the schema " & _
           "(Windows 2003)? Trying without it."
    Query = QueryBase & _
            "(&(objectCategory=computer)" & _
            "(userAccountControl:1.2.840.113556.1.4.803:=8192));" & _
            "dNSHostName;subtree"

    LogStr "A:Getting list of all DCs: " & Query
    Set rs = conn.Execute(Query)
    LogErrorIfNeeded "Error executing query"

  End If

  Dim DCs(), DCCount
  DCCount = 0
  While Not rs.EOF
    LogStr "A:Found " & rs.Fields(0).Value
    ReDim Preserve DCs(DCCount)
    DCs(DCCount) = rs.Fields(0).Value
    LogErrorIfNeeded "Error getting DC name"
    rs.MoveNext
    DCCount = DCCount + 1
  Wend
  GetWritableDCs = DCs
End Function

' This function is based on the sample from
' http://www.vbsedit.com/scripts/desktop/info/scr_231.asp
Function CheckComputerAndUserDetails()
  On Error Resume Next

  Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
  LogErrorIfNeeded "Getting WMI service object for computer details"
  Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
  LogErrorIfNeeded "Executing WMI query for computer details"
  For Each objItem In colItems
    LogStr "A:Computer Name: " & objItem.Name
    LogStr "A:Computer's Domain: " & objItem.Domain
    LogStr "A:Part Of Domain: " & objItem.PartOfDomain
    Select Case objItem.DomainRole
      Case 0 strDomainRole = "Standalone Workstation"
      Case 1 strDomainRole = "Member Workstation"
      Case 2 strDomainRole = "Standalone Server"
      Case 3 strDomainRole = "Member Server"
      Case 4 strDomainRole = "Backup Domain Controller"
      Case 5 strDomainRole = "Primary Domain Controller"
      Case Else strDomainRole = "Unknown (" & objItem.DomainRole & ")"
    End Select
    LogStr "A:Computer's Domain Role: " & strDomainRole
    LogStr "A:Computer's Roles: " & Join(objItem.Roles, ", ")

    If Not objItem.PartOfDomain Then
      LogStr "E:This machine isn't part of a domain. Exiting."
      ErrorMsgBox "This machine isn't part of a domain. Make sure you " & _
                  "are logged in as a domain admin, and run this tool again."
      CheckComputerAndUserDetails = False
      Exit Function
    End If
    UserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
    LogStr "A:Current user's name: " & UserName
    UserDNSDomain = LCase(objShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%"))
    If UserDNSDomain = "%userdnsdomain%" Then
      LogStr "E:The logged in user isn't a domain user. Exiting."
      ErrorMsgBox "The logged in user (" & UserName & ") isn't a domain " & _
                  "user. Make sure you are logged in as a domain admin, " & _
                  "and run this tool again."
      CheckComputerAndUserDetails = False
      Exit Function
    End If
    LogStr "A:Current user's AD DNS domain: " & UserDNSDomain
    If LCase(objItem.Domain) <> UserDNSDomain Then
      LogStr "E:The user's domain doesn't match the machine's domain. Exiting."
      ErrorMsgBox "The current user's DNS domain (" & UserDNSDomain & _
                  ") doesn't match the machine's DNS domain (" & _
                  objItem.Domain & "). This will cause G Suite " & _
                  "Password Sync to fail. Make sure you are logged in as a " & _
                  "domain admin from the same domain as the Domain " & _
                  "Controller, and try the installation again."
      CheckComputerAndUserDetails = False
      Exit Function
    End If
  Next
  CheckComputerAndUserDetails = True
End Function

' This function is based on
' http://www.aspfree.com/c/a/Windows-Scripting/Compressed-Folders-in-WSH/
Function CompressFolder(strPath, strFolder)
  On Error Resume Next

  Const adTypeBinary = 1
  Const adTypeText = 2
  Const adSaveCreateNotExist = 1
  Const adSaveCreateOverwrite = 2
  With CreateObject("ADODB.Stream")
    .Open
    LogErrorIfNeeded "Error opening ADODB for creating the ZIP file"
    .Type = adTypeText
    .WriteText ChrB(&h50) & ChrB(&h4B) & ChrB(&h5) & ChrB(&h6)
    For i = 1 To 18
      .WriteText ChrB(&h0)
    Next
    .SaveToFile strPath, adSaveCreateNotExist
    LogErrorIfNeeded "Error saving ZIP file"
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
    LogErrorIfNeeded "Error re-saving ZIP file"
  End With
  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.NameSpace(strPath)
  LogErrorIfNeeded "Error opening ZIP file for writing"
  intCount = objFolder.Items.Count
  objFolder.CopyHere strFolder, 256
  LogErrorIfNeeded "Error copying files to ZIP file"
  Do Until objFolder.Items.Count = intCount + 1
    WScript.Sleep 200
  Loop
End Function

' All functions below either taken from or based on
' http://explodingcoder.com/blog/content/how-query-active-directory-security-group-membership
' Shawn Poulson, 2009.05.18
' explodingcoder.com

' Returns True if the current user is a Domain Admin, otherwise False
Function CheckIfRunningAsDomainAdmin()
  ' Written by Liron Newman based on Shawn Poulson's example
  ' NOTE: This function doesn't take into account the actual token's groups,
  ' meaning that if running unelevated on a system that uses UAC, the script
  ' will not be able to actually use all the user's permissions.
  On Error Resume Next

  Set oADSysInfo = CreateObject("ADSystemInfo")
  LogErrorIfNeeded "Error creating ADSystemInfo object"
  userDN = oADSysInfo.UserName  ' Get DN of user
  LogStr "A:Current user DN: " & userDN
  ' We shouldn't use the name "Domain Admins" to check membership because it
  ' may be localized, we should use the Well-Known SID.

  ' Define the Domain Admins group SID prefix and suffix in hex:
  Const DomainAdminsSIDStart = "010500000000000515000000"
  Const DomainAdminsSIDEnd = "00020000"
  ' Enumerate all member group names
  tkUser = GetTokenGroups(userDN) ' Get tokens of member groups
  ' See if the Domain Admins group SID is in the token groups
  CheckIfRunningAsDomainAdmin = False
  For Each sid In tkUser
    Dim tmpstr
    tmpstr = ByteArrToHexString(sid)
    If (Left(tmpstr, Len(DomainAdminsSIDStart)) = DomainAdminsSIDStart) _
        And (Right(tmpstr, Len(DomainAdminsSIDEnd)) = DomainAdminsSIDEnd) Then
      CheckIfRunningAsDomainAdmin = True
      Exit For
    End If
    LogErrorIfNeeded "Error checking SID " & tmpstr
  Next
  If CheckIfRunningAsDomainAdmin Then
    LogStr "A:The current user is a member of Domain Admins"
  Else
    LogStr "E:The current user is *not* a member of Domain Admins"
    ErrorMsgBox "The current user isn't a member of the Domain Admins " & _
                "group. To successfully install and setup G Suite " & _
                "Password Sync, you must be a Domain Admin." & _
                vbNewLine & vbNewLine & _
                "Please contact a Domain Admin to continue. You can try " & _
                "running this command, it may add you to the Domain Admins " & _
                "group:" & vbNewLine & vbNewLine & _
                "net group ""Domain Admins"" " & _
                objShell.ExpandEnvironmentStrings("%username%") & " /add" & _
                vbNewLine & vbNewLine & _
                "After joining the Domain Admins group, log out and back " & _
                "in, and try again."
    ' TODO: Get the correct sAMAccountName for Domain Admins, as it may have
    ' been localized... It can be obtained using:
    ' GetObject("LDAP://<SID=" & ByteArrToHexString(objectSid) & ">").Get("sAMAccountName")
  End If
End Function

' Gets tokenGroups attribute from the provided DN
' Usage: <Array of tokens> = GetTokenGroups(<DN of object>)
Function GetTokenGroups(dnObject)
  Dim adsObject

  ' Setup query of tokenGroup SIDs from dnObject
  Set adsObject = GetObject("LDAP://" & Replace(dnObject, "/", "\/"))
  LogErrorIfNeeded "Error opening admin's DN using ADSI"
  adsObject.GetInfoEx Array("tokenGroups"), 0
  GetTokenGroups = adsObject.GetEx("tokenGroups")
  LogErrorIfNeeded "Error getting current user's tokenGroups"
End Function

' Encode Byte() to hex string
Function ByteArrToHexString(bytes)
   Dim i
   ByteArrToHexString = ""
   For i = 1 to Lenb(bytes)
      ByteArrToHexString = ByteArrToHexString & _
                           Right("0" & Hex(Ascb(Midb(bytes, i, 1))), 2)
      LogErrorIfNeeded "Error converting SID bytes to string at " & _
                       ByteArrToHexString
   Next
End Function

' Plans For the future:
' Support non-English systems (i.e. where folder paths are not in English).
' Make an HTML report instead of just collecting text files
' Support paths on upgraded systems such as "C:\WINNT\Profiles\All Users\Application Data\\Google\Google Apps Password Sync\config.xml" etc.
' Offer to restart DCs whose DLL is registered but not loaded, if not the current server
' Offer to start the service is it's stopped
' Ask which username is affected and get their LDIF dump, and correlate their pwdLastSet to appearance in the logs - this can tell us if the issue is with the service, the DLL, etc.
' Get the user's LDIF dump using the credentials detailed in the XML
' Try to find password change events in the Event Log for that user to see where password change occurred
' If any of the log files are missing, collect the ACL of that folder (in case there are no permissions for the service user to create the logs). Offer to fix.
' Ask what user was used to install on the other DCs so that we can get the correct path for UI logs, instead of guessing
' Compare XMLs across all servers
' Get relevant events from Windows Event Logs using "wevtutil"
' Get minidump files: %temp%\WER* folder on Win2008, C:\WINDOWS\pchealth\ERRORREP\UserDumps on Win2003
' Check certificates using certutil -store \\SERVERNAME\AuthRoot | find "Equifax" (or something similar)
' Compare time across DCs
' Compare time to google.com
