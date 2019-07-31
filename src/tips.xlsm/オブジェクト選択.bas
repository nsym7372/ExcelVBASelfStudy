Attribute VB_Name = "オブジェクト選択"
Option Explicit

' ワークブック、ワークシートの明示
Sub ExplicitObject()
    Debug.Print Worksheets(1).Name ' アクティブなブックにより、対象が変化
    Debug.Print ThisWorkbook.Worksheets(1).Name ' マクロが存在する、このブックを明示
    Debug.Print ThisWorkbook.Worksheets(Sheet1.Name).Name   'シート名をオブジェクト名で指定
    Debug.Print Sheet1.Name ' ActiveBookには影響されない模様（上記と同義）
End Sub

' ワークシートをオブジェクト名で指定
Sub SelectSheet()
    Worksheets(Sheet3.Name).Range("A1").Value = "Hello World!"
End Sub

' セル選択
Sub Selection()
    Range("A1").Select ' セル一つ
    Range("A1:C3").Select ' 範囲
    Range("A1", "C3").Select ' 上と同じ
    Range(Cells(1, 1), Cells(3, 3)).Select ' これでもOK、上と同じ
    Range("CellName").Select ' セル名称を指定
    
    Cells(1, 3).Select ' C1セルを選択
    Cells(1, "C").Select ' 上と同じ

End Sub


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

' 範囲内の数式が入ったセルを選択
Sub SpecalizedCell()
    '　「条件を指定してジャンプ」で選択できるものに相当
     Range("B5:F10").SpecialCells(xlCellTypeFormulas).Select
End Sub
