Attribute VB_Name = "選択"
Option Explicit

' 最終行を選択
Sub SelectLastRow()
    Worksheets(1).Cells(Rows.Count, Range("A1").Columns.Count).End(xlUp).Select
End Sub

' 結合セルの情報取得
Sub MergedCellDetail()
    
    Dim icol, irow, icell
    Dim mc As Range
    With Worksheets(1).Range("MergeAreaに含まれるセル")
    
'        指定したセルを含む、結合セルの個数
        icol = .MergeArea.Columns.Count
        irow = .MergeArea.Rows.Count
        icell = .MergeArea.Cells.Count

'        指定したセルそのものの行、列インデックス
        icol = .Column
        irow = .Row
        ' icell = .cell 　プロパティなし、Rangeそのもの

    End With
    
End Sub
