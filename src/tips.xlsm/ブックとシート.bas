Attribute VB_Name = "�u�b�N�ƃV�[�g"
Option Explicit

Sub AddSheet()
  
'   �V�[�g��ǉ�
    Dim sheet As Worksheet
    Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    
'   �ǉ������V�[�g�ɒl���R�s�[
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Call ws.Range("A1:A10").Copy(sheet.Range("A1"))
    
End Sub

Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   �V�����u�b�N�ɃV�[�g��ǉ����āA�l���R�s�[
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
 
End Sub
