Attribute VB_Name = "ファイル"
Option Explicit

Sub ReadWrite()
    Dim theBook As Workbook
    Set theBook = Workbooks.Open(ThisWorkbook.Path & "\sample.xlsx")
    theBook.Worksheets(1).Range("A1").Value = "Hello World"
    
    theBook.Close savechanges:=True
    
    ' 保存
    ' theBook.Save
    
    ' 別名で保存
    ' theBook.SaveAs (ThisWorkbook.Path & "\copy.xlsx")
    
End Sub
