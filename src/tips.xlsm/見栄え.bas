Attribute VB_Name = "見栄え"
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
