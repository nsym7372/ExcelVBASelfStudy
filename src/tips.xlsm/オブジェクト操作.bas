Attribute VB_Name = "オブジェクト操作"
Option Explicit

' 幅を自動調整して、ちょっと余白を加える
Sub ColWidth()
    
    Dim rng As Range
    With Range("C:E")
        .EntireColumn.AutoFit
        For Each rng In .Columns
            rng.ColumnWidth = rng.ColumnWidth + 2
        Next
    End With

End Sub

' セル結合とClear
Sub MergeAndClear()
    ' 以下二つは同じ結果
    Range("A1").Value = "Hello world"
    Range("A1:B2").Merge
    
    Range("C1").Value = "Hello world"
    Range("C1:D2").Merge

'    Range("A1").ClearContents  結合範囲を全て選択しないとエラー
    Range("A1").MergeArea.ClearContents 'OK
    Range("C1").Value = ""  '見かけ上は同じだが、値は存在する
End Sub

' ワークブックにシートを追加
Sub AddSheet()
    '   シートを追加
        Dim sheet As Worksheet
        Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
End Sub

