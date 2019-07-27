Attribute VB_Name = "ブックとシート"
Option Explicit

Sub AddSheet()
  
'   シートを追加
    Dim sheet As Worksheet
    Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    
'   追加したシートに値をコピー
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Call ws.Range("A1:A10").Copy(sheet.Range("A1"))
    
End Sub

Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   新しいブックにシートを追加して、値をコピー
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
 
End Sub
