Attribute VB_Name = "�I�u�W�F�N�g�I��"
Option Explicit

' ���[�N�u�b�N�A���[�N�V�[�g�̖���
Sub ExplicitObject()
    Debug.Print Worksheets(1).Name ' �A�N�e�B�u�ȃu�b�N�ɂ��A�Ώۂ��ω�
    Debug.Print ThisWorkbook.Worksheets(1).Name ' �}�N�������݂���A���̃u�b�N�𖾎�
    Debug.Print ThisWorkbook.Worksheets(Sheet1.Name).Name   '�V�[�g�����I�u�W�F�N�g���Ŏw��
    Debug.Print Sheet1.Name ' ActiveBook�ɂ͉e������Ȃ��͗l�i��L�Ɠ��`�j
End Sub

' ���[�N�V�[�g���I�u�W�F�N�g���Ŏw��
Sub SelectSheet()
    Worksheets(Sheet3.Name).Range("A1").Value = "Hello World!"
End Sub

' �Z���I��
Sub Selection()
    Range("A1").Select ' �Z�����
    Range("A1:C3").Select ' �͈�
    Range("A1", "C3").Select ' ��Ɠ���
    Range(Cells(1, 1), Cells(3, 3)).Select ' ����ł�OK�A��Ɠ���
    Range("CellName").Select ' �Z�����̂��w��
    
    Cells(1, 3).Select ' C1�Z����I��
    Cells(1, "C").Select ' ��Ɠ���

End Sub


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

' �͈͓��̐������������Z����I��
Sub SpecalizedCell()
    '�@�u�������w�肵�ăW�����v�v�őI���ł�����̂ɑ���
     Range("B5:F10").SpecialCells(xlCellTypeFormulas).Select
End Sub
