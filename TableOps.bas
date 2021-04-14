Attribute VB_Name = "TableOps"
Option Explicit


Option Explicit

Sub splitTableByEachRow()
Dim r As Long, cel As Cell, s As Long, e As Long, s1 As Long, e1 As Long, rng As Range
Dim inlsp As InlineShape
r = 1

With Selection
    Set rng = .Range
    Do While (.Information(wdWithInTable))
        .SplitTable
        Set cel = Selection.Document.Tables(r).Cell(1, 8)
        If cel.Range.InlineShapes.Count > 0 Then
        Else
            If Selection.Document.Tables(r).Rows.Count > 1 Then _
                Set cel = Selection.Document.Tables(r).Cell(2, 8)
        End If
        s = .Start: e = .End
        rng.SetRange s, s
        If cel.Range.InlineShapes.Count > 0 Then
             cel.Range.InlineShapes(1).Select
            .Cut
'            cel.Range.InlineShapes(1).Range.Cut ' �Y�n��Range�h�O�o�nDoEvents���t�ΰŶKï�����u�@
'            DoEvents'�γ\��}�l�٦�A�[�F�٬O�|�X���C�٬O��Selection����~�O�I�B�U�L�@��
            s1 = .Start: e1 = .End
            If s1 > s Then
                Do While (rng.Information(wdWithInTable))
                    s1 = s1 - 1
                    rng.SetRange s1, s1
                Loop
            ElseIf s1 < s Then
                Do While (rng.Information(wdWithInTable))
                    s1 = s1 + 1
                    rng.SetRange s1, s1
                Loop
            End If
            rng.Select
            .Paste
            If .Previous.InlineShapes.Count > 0 Then
                With .Previous.InlineShapes(1)
                    .LockAspectRatio = msoTrue
                    .Height = 200
                End With
            Else
                .MoveRight wdCharacter, 1, wdExtend
                With .InlineShapes(1)
                    '.LockAspectRatio = msoTrue
                    .Height = .Height + 181
                    .Width = .Width + 181
                End With
            End If
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.Document.Tables(r).Columns(8).Cells.Delete
        End If
        r = r + 1
        If Selection.Document.Tables(r).Rows.Count > 1 Then
            Selection.Document.Tables(r).Rows(2).Select
        Else
            Exit Do
        End If
    Loop
End With
Beep
End Sub


