Attribute VB_Name = "�I�u�W�F�N�g����"
Option Explicit

' ���������������āA������Ɨ]����������
Sub ColWidth()
    
    Dim rng As Range
    With Range("C:E")
        .EntireColumn.AutoFit
        For Each rng In .Columns
            rng.ColumnWidth = rng.ColumnWidth + 2
        Next
    End With

End Sub

' �Z��������Clear
Sub MergeAndClear()
    ' �ȉ���͓�������
    Range("A1").Value = "Hello world"
    Range("A1:B2").Merge
    
    Range("C1").Value = "Hello world"
    Range("C1:D2").Merge

'    Range("A1").ClearContents  �����͈͂�S�đI�����Ȃ��ƃG���[
    Range("A1").MergeArea.ClearContents 'OK
    Range("C1").Value = ""  '��������͓��������A�l�͑��݂���
End Sub

' ���[�N�u�b�N�ɃV�[�g��ǉ�
Sub AddSheet()
    '   �V�[�g��ǉ�
        Dim sheet As Worksheet
        Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
End Sub

