Attribute VB_Name = "��r�B�z"
Option Explicit
Dim rst As Recordset, d As Object
Dim db As Database 'set db=CurrentDb _
�u��b�w�}�Ҥ�Access���ѷӤ@�� , �G���H�W���ѷ� _
,���HSet db = DBEngine.Workspaces(0).OpenDatabase _
    ("d:\�d�{�@�o�N\���y���\���W.mdb")!���Φ��ѷ�! _
    �Ѧ�: _
    Dim dbsCurrent As Database, dbsContacts As Database'�� CurrentDb ���u�W�����ƻs _
    Set dbsCurrent = CurrentDb _
    Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")


Sub �r�W() '2002/11/10�nSub�~��bWord������!
On Error GoTo ���~�B�z
Dim ch, wrong As Long
'Dim chct As Long
Dim StTime As Date, EndTime As Date
'Dim x As Long, firstword As String '�ýX�ˬd!2002/11/13
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "�r�W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb '�@�w�n�[��d��!!�g���H�U��i!
'�H�W�i�֦��U�G���Y�i!�����|��ܦb����W,�u��@����p���!(��OpenCurrentDatabase���u�W����)
'Set db = d.DBEngine.OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
'Set db = d.DBEngine.Workspaces(0).OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
Set rst = db.OpenRecordset("�r�W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM �r�W��"
End If
StTime = Time
With ActiveDocument
    For Each ch In .Characters '���ýX�r��ch�|�Ǧ^"?"�ܦ��F�B��βŸ�
        wrong = wrong + 1 '�˵���!
'        If wrong = 373 Then MsgBox "Check!!" '�ˬd��!
        If wrong Mod 27250 = 0 Then 'If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
            MsgBox "�]�t�έt���F�췥��,�аȥ�������Access���}��ƪ������,�A�^�ӫ��U�T�w���s�~��!!" _
                , vbExclamation, "���t�έ��n��T��"
'        ElseIf wrong = 49761 Then
'            MsgBox "���ˬd!!"
        End If
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print ch & vbCr & "--------"
        '����r���B�_��r�����p!
'        If Right(ch, 1) <> Chr(10) Or Left(ch, 1) <> Chr(13) Then
        Select Case Asc(ch)
            Case Is <> 13, 10
        With rst
11          .FindFirst "�r�J like '" & ch & "'"
12          If .NoMatch Then
                .AddNew
                rst("�r�J") = ch
                rst("����") = 1
                rst("Asc") = Asc(ch)
                rst("AscW") = AscW(ch)
    '            On Error GoTo ����
                .Update
            Else '���ýX�r��,�|��������B�⤸"?"(Asc(ch)=63),�h�i��b��󤤲Ĥ@���X�{���r�|�~�W����
                '���~�p"�b"�r��(�bWord�����J���Ÿ����̫�@��)�r,��|�P�P�Φr�P�r���X(Asc), _
                ���b�Ÿ����o�����P��m,�N���P�r!�b�έp��,�t�Υ�|�~��b�@�_! _
                �o�I�ٶ��n�J�A!2002/11/13���ծ�,���ɤS�|���}!(��Asc�h�ۦP!)
'                If .AbsolutePosition < 1 And ch Like "?" And Not rst("�r�J") = "?" Then
'                    'If x = 1 Then MsgBox "���ýX�r,���ƱN�[�J�Ĥ@�ӥX�{���r��!!"
'                    MsgBox "���ýX�r,���ƱN�[�J�Ĥ@�ӥX�{���r��!!"
'                    AppActivate "Microsoft Word"
'                    Selection.Collapse
'                    Selection.SetRange wrong + ActiveDocument.Paragraphs.Count / 2, wrong + 1 '�N�ӶýX�r���
'                    x = x + 1
'                End If
                .Edit
                rst("����") = rst("����") + 1
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
'        firstword = "�����ýX�r�[�J�Ĥ@�r:�u" & rst("�r�J") & "�v���@��" & x & "��!!"
'    Else
'        firstword = "����ߧa!�ýX�r��έp���T!!��"
'    End If
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count & vbCr '_
'        & firstword
'    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
'        & vbCr & "���Ӯ�:" & DateDiff("n", StTime, EndTime) & "������" _
'        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "�r�W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number
    Case Is = 91, 3078 '�ѷӤ���DataBase�������
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
'        d.CurrentDb.Close
'        Set db = DBEngine.Workspaces(0).OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
''        Debug.Print Err.Description '�ˬd��!
'        Resume
'    Case Is = 3163 '����r���B�_��r�����p!
'        If Right(ch, 1) = Chr(10) Then
'            ch = Left(ch, Len(ch) - 1)
'        ElseIf Left(ch, 1) = Chr(13) Then
'            ch = Right(ch, Len(ch) - 1) '��If Asc(ch)=13
'        End If
'        Resume 11
    Case Is = 93 '��[]���B�⦡�S��r���ҳ]�������
        rst.FindFirst "asc(�r�J) = " & Asc(ch)
        Resume 12
'    Case Is = -2147023170
'        MsgBox Err.Number & ":" & Err.Description
'        MsgBox Err.LastDllError & "." & Err.Source
'        Set d = CreateObject("access.application")
'        d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
'        d.UserControl = True
'        Resume
'    Case Is = 462 '"���ݦ��A�����s�b�εL�k�ϥ�"
'        'd.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
''        Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
'        Set db = d.CurrentDb
'        Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
'        Resume
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���W() '2002/11/10
On Error GoTo ���~�B�z
Dim Wd, wrong As Long
Dim wrongmark As Integer ', wdct As Long
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True '�p�G��False�hdb.close�|������Ʈw!
'd.UserControl = False
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��UserControl=True�h�����Ϸ|�P�~!
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then db.Execute "DELETE * FROM ���W��"
StTime = Time
With ActiveDocument
    For Each Wd In .Words
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print wd & vbCr & "--------"
        If Len(Wd) > 1 And Right(Wd, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo retry '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        rst.FindFirst "���J like '" & Wd & "'"
        If rst.NoMatch Then
            rst.AddNew
            rst("���J") = Wd
'            On Error GoTo ����
            rst.Update
        Else
            rst.Edit
            rst("����") = rst("����") + 1
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
MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
    & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
    & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��"
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
'����:
'    wrongmark = Err.Number
''    Err.Description = wd
'    If wrongmark = 3022 Then '���ƤF
''        wrong = wrong + 1
''        rst.Seek "=", "���J"
'        rst.FindFirst "���J like '" & wd & "'"
'        rst.Edit
'        rst("����") = rst("����") + 1
'        rst.Update
'        Resume retry
'    Else
'        MsgBox "�����~,���ˬd!!" & Err.Description, vbExclamation
'    End If
End Sub
Sub �i�����W() '2002/11/10�nSub�~��bWord������!'2005/4/21���k�b�]�j�ɮ׮ɤӨS�Ĳv�F!!�]�F3��3�]300��������ɨ�1-3�r���]����!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As Byte 'As String
Dim Dw As String, dwL As Long
length = InputBox("�Ы��w���R���J���W��,�̦h���Ӧr", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 5 Then End
Options.SaveInterval = 0 '�����۰��x�s
StTime = Time
Set d = CreateObject("access.application")
'��Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
'With ActiveDocument
With ActiveDocument
    Dw = .Content '��󤺮e
    dwL = Len(Dw) '������
    .Close
End With
    For phralh = 1 To length 'CByte(length)
'    For phralh = 1 To 5 '�ȩw�̪���5�Ӧr�c������(���i��@�ܼ�)
        For phra = 1 To dwL '.Characters.Count
            Select Case phralh
                Case Is = 1
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    phras = .Characters(phra)'���k�ӺC!
                    phras = Mid(Dw, phra, 1)
                Case Is = 2
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 1 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1)
                    If phra + 1 <= dwL Then phras = Mid(Dw, phra, 2)
                Case Is = 3
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 2 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2)
                    If phra + 2 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 4
                    On Error GoTo ���~�B�z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 3 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3)
                    If phra + 3 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 5
                    On Error GoTo ���~�B�z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 4 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4)
                    If phra + 4 <= dwL Then phras = Mid(Dw, phra, 3)
            End Select
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '�p��
                GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
            End If
            '�����i�J�U�@�Ӧr����
            wrong = wrong + 1 '�˵���!
            If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
                DoEvents 'MsgBox "���ˬd!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "���ˬd!!"
            End If
'            if rst Set rst = CurrentDb.OpenRecordset("SELECT  ���W��.* FROM ���W�� WHERE (((���W��.���J) like '" & phras & "'));")
            With rst
'                If .RecordCount = 0 Then
                .FindFirst "���J like '" & phras & "'"
                If .NoMatch Then
'                    .MoveLast
                    .AddNew
                    rst("���J") = phras
'                    rst("����") = 1'�w�]�Ȥw��1
                    On Error GoTo ���~�B�z
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("����") = rst("����") + 1
                    .Update
                End If
'                .Close
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & dwL '.Characters.Count
'End With
'd.Visible = True
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access'2002/11/15
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 3022
        rst.Requery
        rst.FindFirst "���J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '���X�����������s�b(���W�L������!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub �i�����W1() '2002/11/15�nSub�~��bWord������!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As String
Dim i As Byte, j As Byte
length = InputBox("�Ы��w���R���J���W��,�̦h255�Ӧr", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 255 Then End
Options.SaveInterval = 0 '�����۰��x�s
StTime = Time
Set d = CreateObject("access.application")
'��Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
j = CByte(length)
With ActiveDocument
    For phralh = 1 To j
'    ��ȩw�̪���5�Ӧr�c������,����@�ܼ�j,�h����Byte�j�p��!
        For phra = 1 To .Characters.Count
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To phralh - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '�p��
                GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
            End If
            '�����i�J�U�@�Ӧr����
            wrong = wrong + 1 '�˵���!
            If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
                MsgBox "���ˬd!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "���ˬd!!"
            End If
            With rst
                .FindFirst "���J like '" & phras & "'"
                If .NoMatch Then
    '                .MoveLast
                    .AddNew
                    rst("���J") = phras
                    rst("����") = 1
                    On Error GoTo ���~�B�z
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("����") = rst("����") + 1
                    .Update
                End If
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
'd.Visible = True
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 3022
        rst.Requery
        rst.FindFirst "���J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '���X�����������s�b(���W�L������!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W() '2002/11/11
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w11�r�Ƶ��W()     '2002/11/15'�H������,�i�@���w�����w�r�ƪ��U�ӵ{��(���Ҭ�11�Ӧr���d��)
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w10�r�Ƶ��W() '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w9�r�Ƶ��W()  '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub


Sub ���w8�r�Ƶ��W()   '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w6�r�Ƶ��W()    '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 5 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w5�r�Ƶ��W()     '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 4 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w4�r�Ƶ��W()       '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 3 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w3�r�Ƶ��W()      '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 2 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w2�r�Ƶ��W()       '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 1 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w1�r�Ƶ��W()        '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
            phras = .Characters(phra)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w7�r�Ƶ��W()      '2002/11/15'�H������,�i�@���w�����w�r�ƪ��U�ӵ{��(���Ҭ�7�Ӧr���d��)
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W1() '2002/11/15'�į���C!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim a1, i As Byte, j As Byte
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u255�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W2() '2002/11/15�į�P��]�p�t���h,���i�ܼƤ�!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim i As Byte, j As Byte
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u255�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.DoCmd.SelectObject acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
j = CByte(phralh)
With ActiveDocument
    For phra = 1 To .Characters.Count
'        If j > 1 Then'�Y�ϬO��r�]�������O�B�z�F!!
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
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.DoCmd.OpenTable "���W��", , acReadOnly
    d.DoCmd.Maximize
End If
d.DoCmd.OpenTable "���W��", , acReadOnly
d.DoCmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���r�W_old()
Dim DR As Range, d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ExcelSheet  As Object, _
    ds As Date, de As Date '
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\"
Set d = ActiveDocument
xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""�r�W.XLSX""�r��,�s��ୱ�W", "�r�W�լd", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "�r�W" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If Not charText = Chr(13) And charText <> "-" And Not charText Like "[a-zA-Z0-9��-��]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If InStr(ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub �r�W�[�@
                        End If
                    'End If
                Else
                    GoSub �r�W�[�@
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If U = 0 Then U = 1 '�Y�L����u�r�W�[�@:�v�Ƶ{��,�Y�L�W�L1�����r�W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(U) As String
Set ExcelSheet = CreateObject("Excel.Sheet")
With ExcelSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = U To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "�r�W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
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
    Doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P���r�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '�μƦr�ۤ�
'    For k = 0 To U 'xT�}�C���C�Ӥ������Pj��
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "�r�W=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'����|�Ϯ���
'Set d = Nothing
de = VBA.Timer
MsgBox "�����I" & vbCr & vbCr & "�O��" & Left(de - ds, 5) & "��!"
ExcelSheet.Application.Visible = True
ExcelSheet.Application.UserControl = True
ExcelSheet.SaveAs xlsp '"C:\Macros\�u�uTEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
'Doc.SaveAs "c:\test1.doc"
AppActivate "microsoft excel"
Exit Sub
�r�W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '�O�U�̰��r�W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
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

Function lEnglish() '�^��j�g�r��
Dim Wd, wdct As Long, i As Byte
For i = 65 To 90
    Debug.Print Chr(i) & vbCr
Next
End Function
Function sEnglish() '�^��p�g�r��
Dim i As Byte
For i = 97 To 122
    Debug.Print Chr(i) & vbCr
Next
End Function

Function Symbol() '���I�Ÿ���
Dim f As Variant
f = Array("�C", "�v", Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", Chr(34), ":", ",", ";", _
    "�K�K", "...", "�^", ")", "-")  '���]�w���I�Ÿ��}�C�H�ƥ�
                                'Chr(-24152)�O�u���v,��Asc��Ʀb���(.SelText)�u���v�ɨ��o;Chr(34):�u"�v
End Function

Sub ����q���Ÿ�()
'��1�q���̫�()
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


Sub �y�r�r���ˬd() '�D�ө����ˬd,2004/8/23
Dim ch
For Each ch In ActiveDocument.Characters
'    If AscW(ch) < -1491 Or AscW(ch) > 19968 Then
    If Asc(ch) < -24256 Or (0 > Asc(ch) And Asc(ch) >= -1468) Then
        ch.Select
        ch.Font.Name = "EUDC"
    End If
Next ch
End Sub

Sub �`�}�Ÿ��m��() '2004/10/17
Dim Wd As Range 'As Range 'Words����Y��@��Range����,���u�W����!
'Dim i As Long ' Integer
'�n�����������b��,�o��words�~�ॿ�T�P�_���Ʀr
���μƦr�ഫ���b�μƦr
With Selection '��H������(ActiveDocument),�����H����d���z,���]���ȦӼv�T,�@�o!
    If .Type = wdSelectionIP Then .Document.Select '�p�G�S������d��(�����J�I)�h�B�z������
    If .Document.path = "" Then
        For Each Wd In .Words
            '�n�O�Ʀr�B�e�ᤣ��[�����Ρe�f�~����I
            If Not Wd.Text Like "��" And Not Wd.Text Like "�e" And Not Wd Like "[[]" And Not Wd Like "[]]" Then
                If IsNumeric(Wd) Then
                    If Wd.End = .Document.Content.StoryLength Or Wd.Start = 0 Then GoTo w '��󤧭����t�~�B�z
                    If Not Wd.Previous Like "��" And Not Wd.Previous Like "�e" And Not Wd.Previous Like "[[]" _
                        And Not Wd.Next Like "��" And Not Wd.Next Like "�f" And Not Wd.Next Like "]" Then
w:                      If Wd <= 20 Then 'Arial Unicode MS[����]��"�A����Ʀr"�u���G�Q��!
                            With Wd
                                '����|����Selection���d��,�G������!
'                                .Select 'Words����Y��@��Range����,���u�W����!
                                .Font.Name = "Arial Unicode MS"
                                Wd.Text = ChrW((9312 - 1) + Wd)
                            End With
                        Else '�W�L20�������}��
                            With Wd
                                .Text = "��" & Wd.Text & "��" '�[�A��
                            End With
        '                    MsgBox "���W�L20�������},�������I", vbCritical
        '                    Do Until .Undo(i) = False '�٭쪽�ܤ����٭�]�٭�Ҧ��ʧ@�^
        '                    i = i + 1
        '                    Loop
        '                    StatusBar = "Undo was successful " & i & " times!!" '�b���A�C��ܤ�r�I
        '                    Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        MsgBox "���槹���I", vbInformation
    Else
        MsgBox "����󤣯�ާ@!", vbCritical
    End If
End With
End Sub

Sub ���μƦr�ഫ���b�μƦr() '2004/10/17-�ѹϮѺ޲z�ƻs�Ｖ���즡�Ф��n�A�|�v�T�r��
Dim FNumArray, HNumArray, i As Byte, e As Range
FNumArray = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��")
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

Sub ������b��()
With Selection
    .Range = StrConv(.Range, vbNarrow)
End With
End Sub
Sub ��A����g�W��()
If Selection.Type = wdSelectionIP Then Selection.HomeKey wdStory: Selection.EndKey wdStory, wdExtend
Selection.Text = Replace(Replace(Selection.Text, "�]", "�q"), "�^", "�r")
End Sub


Sub �հɤ�r�Ц�() '2009/8/23
Register_Event_Handler
'���w��F2
' ����2 ����
' �������s�� 2009/8/23�A���s�� Oscar Sun
'
'    Selection.MoveDown Unit:=wdLine, Count:=2
'    Selection.EndKey Unit:=wdLine
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1
'    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionIP Then Exit Sub
    With Selection.Font.Shading
        If InStr(ActiveDocument.Name, "�ƦL") Then
            .Parent.COLOR = wdColorRed
            .Texture = wdTextureNone
        Else
            If .Texture = wdTextureNone Then '�r������
                .Texture = wdTexture15Percent
                .ForegroundPatternColor = wdColorBlack
                .BackgroundPatternColor = wdColorWhite
                .Parent.COLOR = wdColorRed
            Else
                .Texture = wdTextureNone '�r������
                .Parent.COLOR = wdColorAutomatic
            End If
        End If
    End With
    If InStr(ActiveDocument.Name, "�ƦL") Then
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
                        MsgBox "�Ш�s�O���C�I�I", vbExclamation
                        Exit Sub
                    End If
                Next i
                .Cells(r, 1).Activate
                DoEvents
                .ActiveSheet.Paste
                .Cells(r, 2).Value = Selection
                .Cells(r, 2).Font.COLOR = wdColorRed
                If Not Selection Like "*[�����������������������������U�@]*" Then
                    .Cells(r, 5) = Len(Selection)
                ElseIf Selection Like "*�@*" Then
                    .Cells(r, 5) = Len(Selection) - 1
                Else
                    .Cells(r, 5) = 1
                End If
                .ActiveWorkbook.Save
                .Cells(.ActiveCell.Row + 1, .ActiveCell.Column).Activate
            End With
        End With
        ��ЩҦb��m����
        OX.WinActivate "Adobe Reader"
        AppActivate "microsoft word"
    End If
End Sub

Sub ���}�s���e��[��A��()
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
'        Selection.TypeText Text:="�n"
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=2
        'Selection.TypeBackspace
        Selection.TypeText Text:="]"
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop 'While .Find.Execute()
End With
End Sub

Sub �j���޸����x�W�޸�()
Dim a, b, i
a = Array(-24153, -24152, -24155, -24154)  '��,��,��,��
b = Array("�u", "�v", "�y", "�z")

With ActiveDocument.Range.Find
    For i = 0 To 3
        '.Text = a(i)
         '.Replacement.Text = b(i)
         .ClearFormatting
         .Execute Chr(a(i)), , , , , , , , , b(i), wdReplaceAll
    Next i
End With
End Sub


Sub ���r�W()
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\"
Set d = ActiveDocument
xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""�r�W.XLSX""�r��,�s��ୱ�W", "�r�W�լd", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "�r�W" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If InStr("()�G>" & Chr(13) & Chr(9) & Chr(10) & Chr(11) & ChrW(12), charText) = 0 And charText <> "-" And Not charText Like "[a-zA-Z0-9��-��]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If InStr(ChrW(9312) & ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]�����j�i~/�_�X" & Chr(-24152) & Chr(-24153), charText) = 0 Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub �r�W�[�@
                        End If
                    'End If
                Else
                    GoSub �r�W�[�@
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If U = 0 Then U = 1 '�Y�L����u�r�W�[�@:�v�Ƶ{��,�Y�L�W�L1�����r�W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(U) As String
'Set ExcelSheet = CreateObject("Excel.Sheet")
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = U To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "�r�W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
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
    Doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P���r�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '�μƦr�ۤ�
'    For k = 0 To U 'xT�}�C���C�Ӥ������Pj��
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "�r�W=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'����|�Ϯ���
'Set d = Nothing
de = VBA.Timer
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
MsgBox "�����I" & vbCr & vbCr & "�O��" & Left(de - ds, 5) & "��!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp '"C:\Macros\�u�uTEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
'Doc.SaveAs "c:\test1.doc"
'AppActivate "microsoft excel"
Exit Sub
�r�W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '�O�U�̰��r�W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '�\Ū�Ҧ�����s��'����k���ݩʵL�k�ϥΡA�]�����R�O�L�k�b�\Ū���ϥΡC
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

Sub �����W() '�Ѥ��r�W���'2015/11/28
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static Ln
Dim xlsp As String
On Error GoTo ErrH:
Set d = ActiveDocument
'If xlsp = "" Then xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
'If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir
'xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""���W.XLSX""�r��,�s��ୱ�W", "���W�լd", xlsp & Replace(d.Name, ".doc", "") & "���W" & StrConv(Time, vbWide) & ".XLSX")
'If xlsp = "" Then Exit Sub
xlsp = ���o�ୱ���| & "\" & Replace(d.Name, ".doc", "") & "_���W" & StrConv(Time, vbWide) & ".XLSX"
If Ln = "" Then Ln = 1
Ln = InputBox("�Ы��w���J����" & vbCr & vbCr & "�ɮ׷|�s�b�ୱ�W�W��:" & vbCr & vbCr & Replace(d.Name, ".doc", "") & "_���W" & StrConv(Time, vbWide) & ".XLSX" & _
                vbCr & vbCr & "���ɮ�", , Ln + 1)
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
        If Not charText Like "*[-'�@ �C�A�B�F�G�H:,;,�q�r�m�n ''�u�v�y�z�]�^�����H�I�]�^�i�j�X""()<>" _
            & ChrW(9312) & Chr(-24153) & Chr(-24152) & ChrW(8218) & Chr(13) & Chr(10) & Chr(11) & ChrW(12) & Chr(63) & Chr(9) & Chr(-24152) & Chr(-24153) & "�����j�i~/�_�X]*" _
            And Not charText Like "*[a-zA-Z0-9��-��]*" And InStr(charText, ChrW(-243)) = 0 And InStr(charText, Chr(91)) = 0 And InStr(charText, Chr(93)) = 0 Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If Not charText Like "*[" & ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I���a�b]*" Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub ���W�[�@
                        End If
                    'End If
                Else
                    GoSub ���W�[�@
                End If
                preChar = charText
            End If
        End If
    Next
End With
12
Dim Doc As New Document, Xsort() As String, U As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
If U = 0 Then U = 1 '�Y�L����u���W�[�@:�v�Ƶ{��,�Y�L�W�L1�������W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(U) As String
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
Doc.ActiveWindow.Visible = False
If d.ActiveWindow.View.ReadingLayout Then ReadingLayoutB = True: d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
'U = UBound(Xsort)
For j = U To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "���W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) / Ln & "�ӡ^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "���W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) / Ln & "�ӡ^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "���W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
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
    Doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P�����J�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

Doc.ActiveWindow.Visible = True

de = VBA.Timer
Doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
Set d = Nothing ' ActiveDocument.Close wdDoNotSaveChanges

Debug.Print Now

MsgBox "�����I" & vbCr & vbCr & "�O��" & Left(de - ds, 5) & "��!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp
Exit Sub
���W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If U < xT(j) Then U = xT(j) '�O�U�̰����W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '�\Ū�Ҧ�����s��'����k���ݩʵL�k�ϥΡA�]�����R�O�L�k�b�\Ū���ϥΡC
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
    
    Case 91, 5941 '�S���]�w�����ܼƩ� With �϶��ܼ�,���X���һݪ��������s�b
        GoTo 12
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        Resume
        End
    
End Select
End Sub


Sub �ѦW���g�W���ˬd()
Dim s As Long, rng As Range, e, trm As String, ans
Static x() As String, i As Integer
On Error GoTo eH
Do
    Selection.Find.Execute "�q", , , , , , True, wdFindAsk
    Set rng = Selection.Range
    rng.MoveEndUntil "�r"
    trm = Mid(rng, 2)
    
    For Each e In x()
        If StrComp(e, trm) = 0 Then GoTo 1
    Next e
2   ans = MsgBox("�O�_���L�u" & trm & "�v�H" & vbCr & vbCr & vbCr & "�����Ы� NO[�_]", vbExclamation + vbYesNoCancel)
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
    Case 92 '�S���]�w For �j�骺��l�� �}�C�|������
        GoTo 2
End Select
End Sub

Sub �ɶ��b����ഫ() '2017/5/13 �]��YOUKU�PYOUTUBE�ɶ��b��줣�P�ӳ]
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
                    If InStr(Trim(myRng), " ") Then '�p�G��2�ӥH�W�ɶ��b
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
                    Else '�p�G�u��1�Ӯɶ��b
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
