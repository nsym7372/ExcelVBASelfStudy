Attribute VB_Name = "�I��"
Option Explicit

' �ŏI�s�̎��̍s��I��
Sub SelectLastRow()
    Worksheets(1).Cells(Rows.Count, Range("A1").Columns.Count).End(xlUp).Select
End Sub
