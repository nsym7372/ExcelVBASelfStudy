Attribute VB_Name = "���h��"
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
