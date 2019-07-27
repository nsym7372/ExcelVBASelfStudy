Attribute VB_Name = "選択"
Option Explicit

' 最終行の次の行を選択
Sub SelectLastRow()
    Worksheets(1).Cells(Rows.Count, Range("A1").Columns.Count).End(xlUp).Select
End Sub
