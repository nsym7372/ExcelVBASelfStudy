Attribute VB_Name = "�I��"
Option Explicit

' �ŏI�s��I��
Sub SelectLastRow()
    Worksheets(1).Cells(Rows.Count, Range("A1").Columns.Count).End(xlUp).Select
End Sub

' �����Z���̏��擾
Sub MergedCellDetail()
    
    Dim icol, irow, icell
    Dim mc As Range
    With Worksheets(1).Range("MergeArea�Ɋ܂܂��Z��")
    
'        �w�肵���Z�����܂ށA�����Z���̌�
        icol = .MergeArea.Columns.Count
        irow = .MergeArea.Rows.Count
        icell = .MergeArea.Cells.Count

'        �w�肵���Z�����̂��̂̍s�A��C���f�b�N�X
        icol = .Column
        irow = .Row
        ' icell = .cell �@�v���p�e�B�Ȃ��ARange���̂���

    End With
    
End Sub
