Attribute VB_Name = "ƒtƒ@ƒCƒ‹"
Option Explicit

Sub ReadWrite()
    Dim theBook As Workbook
    Set theBook = Workbooks.Open(ThisWorkbook.Path & "\sample.xlsx")
    theBook.Worksheets(1).Range("A1").Value = "Hello World"
    
    theBook.Close savechanges:=True
    
    ' •Û‘¶
    ' theBook.Save
    
    ' •Ê–¼‚Å•Û‘¶
    ' theBook.SaveAs (ThisWorkbook.Path & "\copy.xlsx")
    
End Sub
