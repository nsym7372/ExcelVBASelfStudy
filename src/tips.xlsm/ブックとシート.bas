Attribute VB_Name = "ブックとシート"
Option Explicit

' ワークシートをオブジェクト名で指定
Sub SelectSheet()
    Worksheets(Sheet3.Name).Range("A1").Value = "Hello World!"
End Sub

' ワークブックにシートを追加
Sub AddSheet()
    '   シートを追加
        Dim sheet As Worksheet
        Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
End Sub

' 値のコピー
Sub CopyData()
     
    With ThisWorkbook.Worksheets("Sheet1")
        .Range("A1:A10").Copy (.Range("C3"))
    End With
    
End Sub

' シート作って、作ったシートに値をコピー
Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   新しいブックにシートを追加して、値をコピー
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
End Sub
