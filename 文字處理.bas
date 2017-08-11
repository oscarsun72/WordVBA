Attribute VB_Name = "¤å¦r³B²z"
Option Explicit
Dim rst As Recordset, d As Object
Dim db As Database 'set db=CurrentDb _
¥u¯à¦b¤w¶}±Ò¤§Access¤¤°Ñ·Ó¤@¦¸ , ¤G¦¸¥H¤Wªº°Ñ·Ó _
,¶·¥HSet db = DBEngine.Workspaces(0).OpenDatabase _
    ("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")!ªº§Î¦¡°Ñ·Ó! _
    °Ñ¦Ò: _
    Dim dbsCurrent As Database, dbsContacts As Database'¥Ñ CurrentDb ªº½u¤W»¡©ú½Æ»s _
    Set dbsCurrent = CurrentDb _
    Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")


Sub ¦rÀW() '2002/11/10­nSub¤~¯à¦bWord¤¤°õ¦æ!
On Error GoTo ¿ù»~³B²z
Dim ch, wrong As Long
'Dim chct As Long
Dim StTime As Date, EndTime As Date
'Dim x As Long, firstword As String '¶Ã½XÀË¬d!2002/11/13
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "¦rÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb '¤@©w­n¥[¡©d¡ª!!¼g¦¨¥H¤U¥ç¥i!
'¥H¤W¥i¨Ö¦¨¤U¤G¦¡§Y¥i!¦ý¤£·|Åã¥Ü¦bÀç¹õ¤W,¥u¯à§@¹õ«á­pºâ¥Î!(¨£OpenCurrentDatabaseªº½u¤W»¡©ú)
'Set db = d.DBEngine.OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
'Set db = d.DBEngine.Workspaces(0).OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
Set rst = db.OpenRecordset("¦rÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM ¦rÀWªí"
End If
StTime = Time
With ActiveDocument
    For Each ch In .Characters '¦³¶Ã½X¦r®Éch·|¶Ç¦^"?"ÅÜ¦¨¤F¹Bºâ¥Î²Å¸¹
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong = 373 Then MsgBox "Check!!" 'ÀË¬d¥Î!
        If wrong Mod 27250 = 0 Then 'If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
            MsgBox "¦]¨t²Î­t²ü¹F¨ì·¥­­,½Ð°È¥²¤Á´«¦ÜAccess¥´¶}¸ê®Æªí«áÃö³¬,¦A¦^¨Ó«ö¤U½T©w«ö¶sÄ~Äò!!" _
                , vbExclamation, "¡¹¨t²Î­«­n¸ê°T¡¹"
'        ElseIf wrong = 49761 Then
'            MsgBox "½ÐÀË¬d!!"
        End If
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print ch & vbCr & "--------"
        '´«¦æ¦r¤¸¡B´_¦ì¦r¤¸¤£­p!
'        If Right(ch, 1) <> Chr(10) Or Left(ch, 1) <> Chr(13) Then
        Select Case Asc(ch)
            Case Is <> 13, 10
        With rst
11          .FindFirst "¦r·J like '" & ch & "'"
12          If .NoMatch Then
                .AddNew
                rst("¦r·J") = ch
                rst("¦¸¼Æ") = 1
                rst("Asc") = Asc(ch)
                rst("AscW") = AscW(ch)
    '            On Error GoTo ¦¸¼Æ
                .Update
            Else '·í¦³¶Ã½X¦r®É,·|¦¨¬°¤ñ¸û¹Bºâ¤¸"?"(Asc(ch)=63),«h¥i¯à¦b¤å¥ó¤¤²Ä¤@¦¸¥X²{ªº¦r·|»~¼W¦¸¼Æ
                '¦¹¥~¦p"Åb"¦rµ¥(¦bWord¤¤´¡¤J¡÷²Å¸¹¤º³Ì«á¤@¨Ç)¦r,¥ç·|»P¦P§Î¦r¦P¦r¤¸½X(Asc), _
                ¦ý¦b²Å¸¹ªí¤¤«o¦³¤£¦P¦ì¸m,¥Nªí¤£¦P¦r!¦b²Î­p®É,¨t²Î¥ç·|»~ºâ¦b¤@°_! _
                ³oÂIÁÙ¶·­n§JªA!2002/11/13´ú¸Õ®É,¦³®É¤S·|¤À¶}!(¦ýAsc«h¬Û¦P!)
'                If .AbsolutePosition < 1 And ch Like "?" And Not rst("¦r·J") = "?" Then
'                    'If x = 1 Then MsgBox "¦³¶Ã½X¦r,¦¸¼Æ±N¥[¤J²Ä¤@­Ó¥X²{ªº¦r¤¤!!"
'                    MsgBox "¦³¶Ã½X¦r,¦¸¼Æ±N¥[¤J²Ä¤@­Ó¥X²{ªº¦r¤¤!!"
'                    AppActivate "Microsoft Word"
'                    Selection.Collapse
'                    Selection.SetRange wrong + ActiveDocument.Paragraphs.Count / 2, wrong + 1 '±N¸Ó¶Ã½X¦r¿ï¨ú
'                    x = x + 1
'                End If
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
        End Select
'        chct = .Characters.Count
'        chct = Selection.StoryLength
'        instr(1+
'        .Select
retry:  Next ch
'    rst.Requery
'    rst.MoveFirst
'    If x > 0 Then
'        firstword = "¡·¡·¶Ã½X¦r¥[¤J²Ä¤@¦r:¡u" & rst("¦r·J") & "¡v¤¤¦@¦³" & x & "¦¸!!"
'    Else
'        firstword = "¡¹©ñ¤ß§a!¶Ã½X¦r¥ç²Î­p¥¿½T!!¡¹"
'    End If
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count & vbCr '_
'        & firstword
'    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
'        & vbCr & "¡°¯Ó®É:" & DateDiff("n", StTime, EndTime) & "¤ÀÄÁ¡°" _
'        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "¦rÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number
    Case Is = 91, 3078 '°Ñ·Ó¤£¨ìDataBase¤ºª«¥ó®É
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
'        d.CurrentDb.Close
'        Set db = DBEngine.Workspaces(0).OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
''        Debug.Print Err.Description 'ÀË¬d¥Î!
'        Resume
'    Case Is = 3163 '´«¦æ¦r¤¸¡B´_¦ì¦r¤¸¤£­p!
'        If Right(ch, 1) = Chr(10) Then
'            ch = Left(ch, Len(ch) - 1)
'        ElseIf Left(ch, 1) = Chr(13) Then
'            ch = Right(ch, Len(ch) - 1) '©ÎIf Asc(ch)=13
'        End If
'        Resume 11
    Case Is = 93 '¬°[]µ¥¹Bºâ¦¡¯S®í¦r¤¸©Ò³]¤§¤ñ¸û¦¡
        rst.FindFirst "asc(¦r·J) = " & Asc(ch)
        Resume 12
'    Case Is = -2147023170
'        MsgBox Err.Number & ":" & Err.Description
'        MsgBox Err.LastDllError & "." & Err.Source
'        Set d = CreateObject("access.application")
'        d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
'        d.UserControl = True
'        Resume
'    Case Is = 462 '"»·ºÝ¦øªA¾¹¤£¦s¦b©ÎµLªk¨Ï¥Î"
'        'd.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
''        Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
'        Set db = d.CurrentDb
'        Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
'        Resume
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub µüÀW() '2002/11/10
On Error GoTo ¿ù»~³B²z
Dim Wd, wrong As Long
Dim wrongmark As Integer ', wdct As Long
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True '¦pªG¬°False«hdb.close·|Ãö³¬¸ê®Æ®w!
'd.UserControl = False
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥ÎUserControl=True«h¦³¦¹¤Ï·|­P»~!
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then db.Execute "DELETE * FROM µüÀWªí"
StTime = Time
With ActiveDocument
    For Each Wd In .Words
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print wd & vbCr & "--------"
        If Len(Wd) > 1 And Right(Wd, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo retry '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        rst.FindFirst "µü·J like '" & Wd & "'"
        If rst.NoMatch Then
            rst.AddNew
            rst("µü·J") = Wd
'            On Error GoTo ¦¸¼Æ
            rst.Update
        Else
            rst.Edit
            rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
            rst.Update
        End If
'        wrong = 1
'        wdct = .Words.Count
'        wdct = Selection.StoryLength
'        instr(1+
'        .Select
retry:  Next Wd
End With
EndTime = Time
AppActivate "Microsoft word"
MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
    & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
    & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°"
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
'¦¸¼Æ:
'    wrongmark = Err.Number
''    Err.Description = wd
'    If wrongmark = 3022 Then '­«½Æ¤F
''        wrong = wrong + 1
''        rst.Seek "=", "µü·J"
'        rst.FindFirst "µü·J like '" & wd & "'"
'        rst.Edit
'        rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
'        rst.Update
'        Resume retry
'    Else
'        MsgBox "¦³¿ù»~,½ÐÀË¬d!!" & Err.Description, vbExclamation
'    End If
End Sub
Sub ¶i¶¥µüÀW() '2002/11/10­nSub¤~¯à¦bWord¤¤°õ¦æ!'2005/4/21¦¹ªk¦b¶]¤jÀÉ®×®É¤Ó¨S®Ä²v¤F!!¶]¤F3¤Ñ3©]300­¶ªº¤å¥óÀÉ¨ú1-3¦rµü¶]¤£§¹!
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As Byte 'As String
Dim Dw As String, dwL As Long
length = InputBox("½Ð«ü©w¤ÀªRµü·J¤§¤W­­,³Ì¦h¤­­Ó¦r", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 5 Then End
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
StTime = Time
Set d = CreateObject("access.application")
'©ÎSet d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
'With ActiveDocument
With ActiveDocument
    Dw = .Content '¤å¥ó¤º®e
    dwL = Len(Dw) '¤å¥óªø«×
    .Close
End With
    For phralh = 1 To length 'CByte(length)
'    For phralh = 1 To 5 '¼È©w³Ìªø¬°5­Ó¦rºc¦¨ªºµü(¤´¥i§ï§@ÅÜ¼Æ)
        For phra = 1 To dwL '.Characters.Count
            Select Case phralh
                Case Is = 1
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ¿ù»~³B²z
                    End If
'                    phras = .Characters(phra)'¦¹ªk¤ÓºC!
                    phras = Mid(Dw, phra, 1)
                Case Is = 2
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ¿ù»~³B²z
                    End If
'                    If phra + 1 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1)
                    If phra + 1 <= dwL Then phras = Mid(Dw, phra, 2)
                Case Is = 3
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ¿ù»~³B²z
                    End If
'                    If phra + 2 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2)
                    If phra + 2 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 4
                    On Error GoTo ¿ù»~³B²z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ¿ù»~³B²z
                    End If
'                    If phra + 3 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3)
                    If phra + 3 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 5
                    On Error GoTo ¿ù»~³B²z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ¿ù»~³B²z
                    End If
'                    If phra + 4 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4)
                    If phra + 4 <= dwL Then phras = Mid(Dw, phra, 3)
            End Select
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
            If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
                DoEvents 'MsgBox "½ÐÀË¬d!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "½ÐÀË¬d!!"
            End If
'            if rst Set rst = CurrentDb.OpenRecordset("SELECT  µüÀWªí.* FROM µüÀWªí WHERE (((µüÀWªí.µü·J) like '" & phras & "'));")
            With rst
'                If .RecordCount = 0 Then
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
'                    .MoveLast
                    .AddNew
                    rst("µü·J") = phras
'                    rst("¦¸¼Æ") = 1'¹w³]­È¤w¬°1
                    On Error GoTo ¿ù»~³B²z
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
'                .Close
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & dwL '.Characters.Count
'End With
'd.Visible = True
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access'2002/11/15
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 3022
        rst.Requery
        rst.FindFirst "µü·J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '¶°¦X¤¤ªº¦¨­û¤£¦s¦b(«ü¶W¹L¤å¥óªø«×!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ¶i¶¥µüÀW1() '2002/11/15­nSub¤~¯à¦bWord¤¤°õ¦æ!
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As String
Dim i As Byte, j As Byte
length = InputBox("½Ð«ü©w¤ÀªRµü·J¤§¤W­­,³Ì¦h255­Ó¦r", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 255 Then End
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
StTime = Time
Set d = CreateObject("access.application")
'©ÎSet d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
j = CByte(length)
With ActiveDocument
    For phralh = 1 To j
'    ­ì¼È©w³Ìªø¬°5­Ó¦rºc¦¨ªºµü,¤µ§ï§@ÅÜ¼Æj,«h­­©óByte¤j¤p¦Õ!
        For phra = 1 To .Characters.Count
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To phralh - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
            If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
                MsgBox "½ÐÀË¬d!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "½ÐÀË¬d!!"
            End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
    '                .MoveLast
                    .AddNew
                    rst("µü·J") = phras
                    rst("¦¸¼Æ") = 1
                    On Error GoTo ¿ù»~³B²z
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
'd.Visible = True
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 3022
        rst.Requery
        rst.FindFirst "µü·J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '¶°¦X¤¤ªº¦¨­û¤£¦s¦b(«ü¶W¹L¤å¥óªø«×!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW() '2002/11/11
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        Select Case CByte(phralh)
            Case Is = 1
                phras = .Characters(phra)
            Case Is = 2
                If phra + 1 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1)
            Case Is = 3
                If phra + 2 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2)
            Case Is = 4
                If phra + 3 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3)
            Case Is = 5
                If phra + 4 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4)
            Case Is = 6
                If phra + 5 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5)
            Case Is = 7
                If phra + 6 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6)
            Case Is = 8
                If phra + 7 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7)
            Case Is = 9
                If phra + 8 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8)
            Case Is = 10
                If phra + 9 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9)
            Case Is = 11
                If phra + 10 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9) & _
                        .Characters(phra + 10)
        End Select
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w11¦r¼ÆµüÀW()     '2002/11/15'¥H¦¹¬°¨Ò,¥i§@¬°¹w¥ý­­©w¦r¼Æªº¦U­Óµ{§Ç(¥»¨Ò¬°11­Ó¦rªº¬d¸ß)
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 10 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8) & .Characters(phra + 9) & _
                    .Characters(phra + 10)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w10¦r¼ÆµüÀW() '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 9 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8) & .Characters(phra + 9)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w9¦r¼ÆµüÀW()  '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 8 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub


Sub «ü©w8¦r¼ÆµüÀW()   '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 7 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w6¦r¼ÆµüÀW()    '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 5 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w5¦r¼ÆµüÀW()     '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 4 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w4¦r¼ÆµüÀW()       '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 3 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w3¦r¼ÆµüÀW()      '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 2 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w2¦r¼ÆµüÀW()       '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 1 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w1¦r¼ÆµüÀW()        '2002/11/15
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
            phras = .Characters(phra)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub «ü©w7¦r¼ÆµüÀW()      '2002/11/15'¥H¦¹¬°¨Ò,¥i§@¬°¹w¥ý­­©w¦r¼Æªº¦U­Óµ{§Ç(¥»¨Ò¬°7­Ó¦rªº¬d¸ß)
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 6 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW1() '2002/11/15'®Ä¯à¸ûºC!
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim a1, i As Byte, j As Byte
phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u255¡v!", "«ü©wµü·J¦r¼Æ", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        j = CByte(phralh)
        ReDim a1(1 To j) As String
        If j > 1 Then
            If phra + (phralh - 1) <= .Characters.Count Then
                For j = 1 To j
                    For i = 0 To j - 1
                            a1(j) = a1(j) & .Characters(phra + i)
                    Next i
    '                    Debug.Print a1(j)
                Next j
                phras = a1(j - 1)
            End If
        Else
            phras = .Characters(phra)
        End If
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW2() '2002/11/15®Ä¯à»P­ì³]­p®t¤£¦h,¦ý¥iÅÜ¼Æ¤Æ!
On Error GoTo ¿ù»~³B²z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim i As Byte, j As Byte
phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u255¡v!", "«ü©wµü·J¦r¼Æ", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
d.DoCmd.SelectObject acTable, "µüÀWªí", True
'd.Visible = True 'ÀË¬d¥Î
Set db = d.CurrentDb
Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
'    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
    db.Execute "DELETE * FROM µüÀWªí"
End If
StTime = Time
j = CByte(phralh)
With ActiveDocument
    For phra = 1 To .Characters.Count
'        If j > 1 Then'§Y¨Ï¬O³æ¦r¤]¤£¶·¤À§O³B²z¤F!!
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To j - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
'        Else
'            phras = .Characters(phra)
'        End If
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '­p¦¸
            GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
        End If
        'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
        wrong = wrong + 1 'ÀËµø¥Î!
'        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
'            MsgBox "½ÐÀË¬d!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "½ÐÀË¬d!!"
'        End If
        With rst
            .FindFirst "µü·J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("µü·J") = phras
'                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
End With
If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "µüÀWªí", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
    Case Is = 91, 3078
        MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ¤å¥ó¦rÀW_old()
Dim DR As Range, d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ExcelSheet  As Object, _
    ds As Date, de As Date '
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\®à­±\"
Set d = ActiveDocument
xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
        "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""¦rÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "¦rÀW½Õ¬d", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If Not charText = Chr(13) And charText <> "-" And Not charText Like "[a-zA-Z0-9¢¯-¢¸]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            If InStr(ChrW(-24153) & ChrW(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            'chr(2)¥i¯à¬Oµù¸}¼Ð°O
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub ¦rÀW¥[¤@
                        End If
                    'End If
                Else
                    GoSub ¦rÀW¥[¤@
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If U = 0 Then U = 1 '­YµL°õ¦æ¡u¦rÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªº¦rÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5

ReDim Xsort(U) As String
Set ExcelSheet = CreateObject("Excel.Sheet")
With ExcelSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = U To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "¡B", Chr(9), 1, 1) 'chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "¦rÀW") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "¼Ð·¢Åé"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªº¦r¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '¥Î¼Æ¦r¬Û¤ñ
'    For k = 0 To U 'xT°}¦C¤¤¨C­Ó¤¸¯À³£»Pj¤ñ
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "¦rÀW=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'¦¹¦æ·|¨Ï®ø¥¢
'Set d = Nothing
de = VBA.Timer
MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & Left(de - ds, 5) & "¬í!"
ExcelSheet.Application.Visible = True
ExcelSheet.Application.UserControl = True
ExcelSheet.SaveAs xlsp '"C:\Macros\¦u¯uTEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
'Doc.SaveAs "c:\test1.doc"
AppActivate "microsoft excel"
Exit Sub
¦rÀW¥[¤@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '°O¤U³Ì°ª¦rÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
        '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
'        Resume
        End
    
End Select
End Sub

Function lEnglish() '­^¤å¤j¼g¦r¥À
Dim Wd, wdct As Long, i As Byte
For i = 65 To 90
    Debug.Print Chr(i) & vbCr
Next
End Function
Function sEnglish() '­^¤å¤p¼g¦r¥À
Dim i As Byte
For i = 97 To 122
    Debug.Print Chr(i) & vbCr
Next
End Function

Function Symbol() '¼ÐÂI²Å¸¹ªí
Dim f As Variant
f = Array("¡C", "¡v", Chr(-24152), "¡G", "¡A", "¡F", _
    "¡B", "¡u", ".", Chr(34), ":", ",", ";", _
    "¡K¡K", "...", "¡^", ")", "-")  '¥ý³]©w¼ÐÂI²Å¸¹°}¦C¥H³Æ¥Î
                                'Chr(-24152)¬O¡u¡¨¡v,¥ÑAsc¨ç¼Æ¦b¿ï¨ú(.SelText)¡u¡¨¡v®É¨ú±o;Chr(34):¡u"¡v
End Function

Sub ¿ï¨ú¬q¸¨²Å¸¹()
'²Ä1¬qªº³Ì«á()
'    With ActiveDocument.Paragraphs(1).Range
'        ActiveDocument.Range(.End - 1, .End).Select
'    End With
Dim i As Integer
For i = 1 To ActiveDocument.Paragraphs.Count
    With ActiveDocument.Paragraphs(i).Range
        ActiveDocument.Range(.End - 1, .End).Select
    End With
Next i
End Sub


Sub ³y¦r¦r¤¸ÀË¬d() '«D²Ó©úÅéÀË¬d,2004/8/23
Dim ch
For Each ch In ActiveDocument.Characters
'    If AscW(ch) < -1491 Or AscW(ch) > 19968 Then
    If Asc(ch) < -24256 Or (0 > Asc(ch) And Asc(ch) >= -1468) Then
        ch.Select
        ch.Font.Name = "EUDC"
    End If
Next ch
End Sub

Sub ª`¸}²Å¸¹¸m´«() '2004/10/17
Dim Wd As Range 'As Range 'Wordsª«¥ó§Yªí¤@­ÓRangeª«¥ó,¨£½u¤W»¡©ú!
'Dim i As Long ' Integer
'­n¥ý°õ¦æ¥þ§ÎÂà¥b§Î,³o¼Ëwords¤~¯à¥¿½T§PÂ_¬°¼Æ¦r
¥þ§Î¼Æ¦rÂà´«¦¨¥b§Î¼Æ¦r
With Selection '­ì¥H¾ã¥÷¤å¥ó(ActiveDocument),¤µ¦ý¥H¿ï¨ú½d³ò¾ã²z,¦ý¦]§ó§ï­È¦Ó¼vÅT,§@¼o!
    If .Type = wdSelectionIP Then .Document.Select '¦pªG¨S¦³¿ï¨ú½d³ò(¬°´¡¤JÂI)«h³B²z¾ã¥÷¤å¥ó
    If .Document.path = "" Then
        For Each Wd In .Words
            '­n¬O¼Æ¦r¥B«e«á¤£¯à¥[¡£¡¤©Î¡e¡f¤~°õ¦æ¡I
            If Not Wd.Text Like "¡£" And Not Wd.Text Like "¡e" And Not Wd Like "[[]" And Not Wd Like "[]]" Then
                If IsNumeric(Wd) Then
                    If Wd.End = .Document.Content.StoryLength Or Wd.Start = 0 Then GoTo w '¤å¥ó¤§­º§À¥t¥~³B²z
                    If Not Wd.Previous Like "¡£" And Not Wd.Previous Like "¡e" And Not Wd.Previous Like "[[]" _
                        And Not Wd.Next Like "¡¤" And Not Wd.Next Like "¡f" And Not Wd.Next Like "]" Then
w:                      If Wd <= 20 Then 'Arial Unicode MS[ºØÃþ]¸Ì"¬A¸¹¤å¼Æ¦r"¥u¦³¤G¤Q­Ó!
                            With Wd
                                '¿ï¨ú·|§ïÅÜSelectionªº½d³ò,¬G¤µ¨ú®ø!
'                                .Select 'Wordsª«¥ó§Yªí¤@­ÓRangeª«¥ó,¨£½u¤W»¡©ú!
                                .Font.Name = "Arial Unicode MS"
                                Wd.Text = ChrW((9312 - 1) + Wd)
                            End With
                        Else '¶W¹L20¸¹ªºµù¸}®É
                            With Wd
                                .Text = "¡£" & Wd.Text & "¡¤" '¥[¬A¸¹
                            End With
        '                    MsgBox "¦³¶W¹L20¸¹ªºµù¸},¤£¯à°õ¦æ¡I", vbCritical
        '                    Do Until .Undo(i) = False 'ÁÙ­ìª½¦Ü¤£¯àÁÙ­ì¡]ÁÙ­ì©Ò¦³°Ê§@¡^
        '                    i = i + 1
        '                    Loop
        '                    StatusBar = "Undo was successful " & i & " times!!" '¦bª¬ºA¦CÅã¥Ü¤å¦r¡I
        '                    Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        MsgBox "°õ¦æ§¹²¦¡I", vbInformation
    Else
        MsgBox "¥»¤å¥ó¤£¯à¾Þ§@!", vbCritical
    End If
End With
End Sub

Sub ¥þ§Î¼Æ¦rÂà´«¦¨¥b§Î¼Æ¦r() '2004/10/17-¥Ñ¹Ï®ÑºÞ²z½Æ»s§ï¼¶ªº­ì¦¡¡Ð¤£¦n¡A·|¼vÅT¦r§Î
Dim FNumArray, HNumArray, i As Byte, e As Range
FNumArray = Array("¢¯", "¢°", "¢±", "¢²", "¢³", "¢´", "¢µ", "¢¶", "¢·", "¢¸")
HNumArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
With ActiveDocument
    For Each e In .Characters
        For i = 1 To UBound(FNumArray) + 1
            If e.Text Like FNumArray(i - 1) Then
                e.Text = HNumArray(i - 1)
        End If
        Next i
    Next e
End With
End Sub

Sub ¥þ§ÎÂà¥b§Î()
With Selection
    .Range = StrConv(.Range, vbNarrow)
End With
End Sub
Sub ¶ê¬A¸¹§ï½g¦W¸¹()
If Selection.Type = wdSelectionIP Then Selection.HomeKey wdStory: Selection.EndKey wdStory, wdExtend
Selection.Text = Replace(Replace(Selection.Text, "¡]", "¡q"), "¡^", "¡r")
End Sub


Sub ®Õ°É¤å¦r¼Ð¦â() '2009/8/23
Register_Event_Handler
'«ü©wÁäF2
' ¥¨¶°2 ¥¨¶°
' ¥¨¶°¿ý»s©ó 2009/8/23¡A¿ý»sªÌ Oscar Sun
'
'    Selection.MoveDown Unit:=wdLine, Count:=2
'    Selection.EndKey Unit:=wdLine
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1
'    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionIP Then Exit Sub
    With Selection.Font.Shading
        If InStr(ActiveDocument.Name, "±Æ¦L") Then
            .Parent.COLOR = wdColorRed
            .Texture = wdTextureNone
        Else
            If .Texture = wdTextureNone Then '¦r¤¸ºô©³
                .Texture = wdTexture15Percent
                .ForegroundPatternColor = wdColorBlack
                .BackgroundPatternColor = wdColorWhite
                .Parent.COLOR = wdColorRed
            Else
                .Texture = wdTextureNone '¦r¤¸ºô©³
                .Parent.COLOR = wdColorAutomatic
            End If
        End If
    End With
    If InStr(ActiveDocument.Name, "±Æ¦L") Then
        ActiveDocument.Save
'        setOX
'        OX.WinActivate "Microsoft Excel"
        Dim e As Excel.Application
        Dim r As Long, i As Byte
        With Selection
            Set e = GetObject(, "Excel.application")
            AppActivate "microsoft excel"
            With e
                '.ActiveWorkbook.Save
                r = .ActiveCell.Row
                For i = 1 To 7
                    If .Cells(r, i).Value <> "" Then
                        MsgBox "½Ð¨ì·s°O¿ý¦C¡I¡I", vbExclamation
                        Exit Sub
                    End If
                Next i
                .Cells(r, 1).Activate
                DoEvents
                .ActiveSheet.Paste
                .Cells(r, 2).Value = Selection
                .Cells(r, 2).Font.COLOR = wdColorRed
                If Not Selection Like "*[óòñõôöø÷ùûúüýþ¡¸¡¹¡U¡@]*" Then
                    .Cells(r, 5) = Len(Selection)
                ElseIf Selection Like "*¡@*" Then
                    .Cells(r, 5) = Len(Selection) - 1
                Else
                    .Cells(r, 5) = 1
                End If
                .ActiveWorkbook.Save
                .Cells(.ActiveCell.Row + 1, .ActiveCell.Column).Activate
            End With
        End With
        ´å¼Ð©Ò¦b¦ì¸m®ÑÅÒ
        OX.WinActivate "Adobe Reader"
        AppActivate "microsoft word"
    End If
End Sub

Sub µù¸}½s¸¹«e«á¥[¤è¬A¸¹()
With Selection
    Do

        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
'        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = False
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchByte = True
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'        If .Find.Execute() = False Then Exit Do
        'Application.Browser.Next
        .TypeText Text:="["
        .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        .Font.Superscript = wdToggle
'        Selection.Copy
'        Selection.MoveRight Unit:=wdCharacter, Count:=3
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Paste
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Delete Unit:=wdCharacter, Count:=1
'        Selection.TypeText Text:="¡n"
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=2
        'Selection.TypeBackspace
        Selection.TypeText Text:="]"
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop 'While .Find.Execute()
End With
End Sub

Sub ¤j³°¤Þ¸¹´«¥xÆW¤Þ¸¹()
Dim a, b, i
a = Array(-24153, -24152, -24155, -24154)  '¡§,¡¨,¡¥,¡¨
b = Array("¡u", "¡v", "¡y", "¡z")

With ActiveDocument.Range.Find
    For i = 0 To 3
        '.Text = a(i)
         '.Replacement.Text = b(i)
         .ClearFormatting
         .Execute Chr(a(i)), , , , , , , , , b(i), wdReplaceAll
    Next i
End With
End Sub


Sub ¤å¥ó¦rÀW()
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\®à­±\"
Set d = ActiveDocument
xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
        "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""¦rÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "¦rÀW½Õ¬d", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If InStr("()¡G>" & Chr(13) & Chr(9) & Chr(10) & Chr(11) & ChrW(12), charText) = 0 And charText <> "-" And Not charText Like "[a-zA-Z0-9¢¯-¢¸]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            If InStr(ChrW(9312) & ChrW(-24153) & ChrW(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]¡¾¡¼¡j¡i~/¡_¡X" & Chr(-24152) & Chr(-24153), charText) = 0 Then
            'chr(2)¥i¯à¬Oµù¸}¼Ð°O
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub ¦rÀW¥[¤@
                        End If
                    'End If
                Else
                    GoSub ¦rÀW¥[¤@
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If U = 0 Then U = 1 '­YµL°õ¦æ¡u¦rÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªº¦rÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5

ReDim Xsort(U) As String
'Set ExcelSheet = CreateObject("Excel.Sheet")
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = U To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "¡B", Chr(9), 1, 1) 'chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "¦rÀW") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "¼Ð·¢Åé"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªº¦r¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '¥Î¼Æ¦r¬Û¤ñ
'    For k = 0 To U 'xT°}¦C¤¤¨C­Ó¤¸¯À³£»Pj¤ñ
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "¦rÀW=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'¦¹¦æ·|¨Ï®ø¥¢
'Set d = Nothing
de = VBA.Timer
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & Left(de - ds, 5) & "¬í!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp '"C:\Macros\¦u¯uTEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
'Doc.SaveAs "c:\test1.doc"
'AppActivate "microsoft excel"
Exit Sub
¦rÀW¥[¤@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '°O¤U³Ì°ª¦rÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
        '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '¾\Åª¼Ò¦¡¤£¯à½s¿è'¦¹¤èªk©ÎÄÝ©ÊµLªk¨Ï¥Î¡A¦]¬°¦¹©R¥OµLªk¦b¾\Åª¤¤¨Ï¥Î¡C
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        'Resume
        End
    
End Select
End Sub

Sub ¤å¥óµüÀW() '¥Ñ¤å¥ó¦rÀW§ï¨Ó'2015/11/28
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static Ln
Dim xlsp As String
On Error GoTo ErrH:
Set d = ActiveDocument
'If xlsp = "" Then xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
'If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir
'xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
        "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""µüÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "µüÀW½Õ¬d", xlsp & Replace(d.Name, ".doc", "") & "µüÀW" & StrConv(Time, vbWide) & ".XLSX")
'If xlsp = "" Then Exit Sub
xlsp = ¨ú±o®à­±¸ô®| & "\" & Replace(d.Name, ".doc", "") & "_µüÀW" & StrConv(Time, vbWide) & ".XLSX"
If Ln = "" Then Ln = 1
Ln = InputBox("½Ð«ü©wµü·Jªø«×" & vbCr & vbCr & "ÀÉ®×·|¦s¦b®à­±¤W¦W¬°:" & vbCr & vbCr & Replace(d.Name, ".doc", "") & "_µüÀW" & StrConv(Time, vbWide) & ".XLSX" & _
                vbCr & vbCr & "ªºÀÉ®×", , Ln + 1)
If Ln = "" Then Exit Sub
If Not IsNumeric(Ln) Then Exit Sub
If Ln > 11 Or Ln < 2 Then Exit Sub


ds = VBA.Timer

With d
    For Each Char In d.Characters
        Select Case Ln
            Case 2
                charText = Char & Char.Next
            Case 3
                charText = Char & Char.Next & Char.Next.Next
            Case 4
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next
            Case 5
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next
            Case 6
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next
            Case 7
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next
            Case 8
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next
            Case 9
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next
            Case 10
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next
            Case 11
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next
        End Select
        If Not charText Like "*[-'¡@ ¡C¡A¡B¡F¡G¡H:,;,¡q¡r¡m¡n ''¡u¡v¡y¡z¡]¡^¡¾¡µ¡H¡I¡]¡^¡i¡j¡X""()<>" _
            & ChrW(9312) & Chr(-24153) & Chr(-24152) & ChrW(8218) & Chr(13) & Chr(10) & Chr(11) & ChrW(12) & Chr(63) & Chr(9) & Chr(-24152) & Chr(-24153) & "¡¾¡¼¡j¡i~/¡_¡X]*" _
            And Not charText Like "*[a-zA-Z0-9¢¯-¢¸]*" And InStr(charText, ChrW(-243)) = 0 And InStr(charText, Chr(91)) = 0 And InStr(charText, Chr(93)) = 0 Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            If Not charText Like "*[" & ChrW(-24153) & ChrW(-24152) & Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I¡¥¡a¡b]*" Then
            'chr(2)¥i¯à¬Oµù¸}¼Ð°O
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub µüÀW¥[¤@
                        End If
                    'End If
                Else
                    GoSub µüÀW¥[¤@
                End If
                preChar = charText
            End If
        End If
    Next
End With
12
Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
If U = 0 Then U = 1 '­YµL°õ¦æ¡uµüÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªºµüÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5

ReDim Xsort(U) As String
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
    Next j
End With
Doc.ActiveWindow.Visible = False
If d.ActiveWindow.View.ReadingLayout Then ReadingLayoutB = True: d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
'U = UBound(Xsort)
For j = U To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                .Range.InsertAfter "µüÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) / Ln & "­Ó¡^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "µüÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) / Ln & "­Ó¡^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "¡B", Chr(9), 1, 1) 'chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "µüÀW") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "¼Ð·¢Åé"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªºµü·J¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
End With

Doc.ActiveWindow.Visible = True

de = VBA.Timer
Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
Set d = Nothing ' ActiveDocument.Close wdDoNotSaveChanges

Debug.Print Now

MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & Left(de - ds, 5) & "¬í!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp
Exit Sub
µüÀW¥[¤@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '°O¤U³Ì°ªµüÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
        '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '¾\Åª¼Ò¦¡¤£¯à½s¿è'¦¹¤èªk©ÎÄÝ©ÊµLªk¨Ï¥Î¡A¦]¬°¦¹©R¥OµLªk¦b¾\Åª¤¤¨Ï¥Î¡C
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = False ' Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    
    Case 91, 5941 '¨S¦³³]©wª«¥óÅÜ¼Æ©Î With °Ï¶ôÅÜ¼Æ,¶°¦X¤¤©Ò»Ýªº¦¨­û¤£¦s¦b
        GoTo 12
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        Resume
        End
    
End Select
End Sub


Sub ®Ñ¦W¸¹½g¦W¸¹ÀË¬d()
Dim s As Long, rng As Range, e, trm As String, ans
Static x() As String, i As Integer
On Error GoTo eH
Do
    Selection.Find.Execute "¡q", , , , , , True, wdFindAsk
    Set rng = Selection.Range
    rng.MoveEndUntil "¡r"
    trm = Mid(rng, 2)
    
    For Each e In x()
        If StrComp(e, trm) = 0 Then GoTo 1
    Next e
2   ans = MsgBox("¬O§_²¤¹L¡u" & trm & "¡v¡H" & vbCr & vbCr & vbCr & "µ²§ô½Ð«ö NO[§_]", vbExclamation + vbYesNoCancel)
    Select Case ans
        Case vbYes
            ReDim Preserve x(i) As String
            x(i) = trm
            i = i + 1
        Case vbNo
            Exit Sub
    End Select
1
Loop
Exit Sub
eH:
Select Case Err.Number
    Case 92 '¨S¦³³]©w For °j°éªºªì©l­È °}¦C©|¥¼¦³­È
        GoTo 2
End Select
End Sub

Sub ®É¶¡¶b³æ¦ìÂà´«() '2017/5/13 ¦]À³YOUKU»PYOUTUBE®É¶¡¶b³æ¦ì¤£¦P¦Ó³]
'Debug.Print Len(ActiveDocument.Range)
Dim a, aM, aMM, s As Long, e As Long
Dim myRng As Range, chRng As Range
Set myRng = ActiveDocument.Range
Set chRng = ActiveDocument.Range
s = -1
For Each a In ActiveDocument.Characters
    If a.Font.Name = "Times New Roman" Then
        If s = -1 Then s = a.Start
        If a = Chr(13) Then GoTo 1
    Else
1       If s > -1 Then
            e = a.Previous.End
            myRng.SetRange s, e
            If InStr(myRng, "http") = 0 Then
                If InStr(Replace(myRng, ":", "", 1, 1), ":") Then 'if find : * 2
                    If InStr(Trim(myRng), " ") Then '¦pªG¦³2­Ó¥H¤W®É¶¡¶b
                        For Each aMM In myRng.Characters
                            If aMM.Next = " " Then
                                e = aMM.End
                                chRng.SetRange s, e
'                                chRng.Select
                                If InStr(Replace(chRng, ":", "", 1, 1), ":") Then 'if find : * 2
                                    GoSub chng
                                End If
                                s = chRng.End + 1
                            End If
                        Next
                    Else '¦pªG¥u¦³1­Ó®É¶¡¶b
                        chRng.SetRange myRng.Start, myRng.End
                        GoSub chng
                    End If
                End If
            End If
            s = -1
        End If
    End If
Next
ActiveDocument.Range.Find.Execute "  ", True, , , , , , wdFindContinue, , " ", wdReplaceAll
Exit Sub
chng:
                    For Each aM In chRng.Characters
                        If aM.Next = ":" Then
                            aM.Next.Next.Text = Str((CInt(aM.Next.Next) * 10 + CInt(aM) * 60) / 10)
                            aM.Next.Delete
                            aM.Delete
                            Exit For
                        End If
                    Next
Return
End Sub
