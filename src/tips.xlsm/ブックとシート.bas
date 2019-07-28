Attribute VB_Name = "�u�b�N�ƃV�[�g"
Option Explicit

' ���[�N�V�[�g���I�u�W�F�N�g���Ŏw��
Sub SelectSheet()
    Worksheets(Sheet3.Name).Range("A1").Value = "Hello World!"
End Sub

' ���[�N�u�b�N�ɃV�[�g��ǉ�
Sub AddSheet()
    '   �V�[�g��ǉ�
        Dim sheet As Worksheet
        Set sheet = Worksheets.Add(after:=Worksheets(Worksheets.Count))
End Sub

' �l�̃R�s�[
Sub CopyData()
     
    With ThisWorkbook.Worksheets("Sheet1")
        .Range("A1:A10").Copy (.Range("C3"))
    End With
    
End Sub

' �V�[�g����āA������V�[�g�ɒl���R�s�[
Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   �V�����u�b�N�ɃV�[�g��ǉ����āA�l���R�s�[
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
End Sub
