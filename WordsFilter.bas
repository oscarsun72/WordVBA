Attribute VB_Name = "WordsFilter"
Option Explicit

Sub �q�����]�t���w�r��()
Dim FilterdWords(), d As Document, p As Paragraph, i As Long
Static pIndex As Long
FilterdWords = Array("github", "visual studio") '��ʿ�J�z�����A����W�L Byte �W��
Set d = ActiveDocument
For Each p In d.Paragraphs
    i = i + 1
    If i > pIndex Then
        If FilterCriteriaAnd(FilterdWords, p) Then
            p.Range.Select
            pIndex = i
            Exit Sub
        End If
    End If
Next
MsgBox "�䧹�F�I", vbInformation
End Sub


Function FilterCriteriaAnd(Strs(), p As Paragraph) As Boolean
Dim uB As Byte, i As Byte
uB = UBound(Strs())
For i = 0 To uB
    If InStr(1, p.Range, Strs(i), vbTextCompare) > 0 Then
        FilterCriteriaAnd = True
    Else
        FilterCriteriaAnd = False
        Exit Function
    End If
Next
End Function

