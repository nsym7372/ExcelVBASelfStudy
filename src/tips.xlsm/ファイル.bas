Attribute VB_Name = "�t�@�C��"
Option Explicit

Sub ReadWrite()
    Dim theBook As Workbook
    Set theBook = Workbooks.Open(ThisWorkbook.Path & "\sample.xlsx")
    theBook.Worksheets(1).Range("A1").Value = "Hello World"
    
    theBook.Close savechanges:=True
    
    ' �ۑ�
    ' theBook.Save
    
    ' �ʖ��ŕۑ�
    ' theBook.SaveAs (ThisWorkbook.Path & "\copy.xlsx")
    
End Sub
