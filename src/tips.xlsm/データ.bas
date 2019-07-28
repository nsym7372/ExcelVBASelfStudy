Attribute VB_Name = "�f�[�^"
Option Explicit

' �z��錾
Sub DeclareArray()

    '�z��錾�ɂ͗v�f�����K�v
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"    '���R�A�^���Ⴄ�ƃG���[
    
    Dim s As Variant
    For Each s In ary
        Debug.Print s
    Next
End Sub

' Collection�錾
Sub DeclareCollection()

    'Collection�ɂ͗v�f���s�v
    Dim sheets As Collection
    Set sheets = New Collection
    
    With sheets
        .Add Sheet1.Name
        .Add Sheet3.Name
        .Add 1  '�^���Ⴆ�ǖ��Ȃ�
    End With
    
    Dim s As Variant
    For Each s In sheets
        Debug.Print s
    Next
End Sub

' For Each
Sub ForEach()
    Dim ary(2) As String
    ary(1) = "Hello"
    ary(2) = "World"
    
    Dim str As Variant 'for each�Ŏg���ꍇ��variant���K�v
    For Each str In ary
        Debug.Print (str)
    Next
End Sub

'�v�f���Ŕ���
Sub ArrayIndex()
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"
    
    Dim i As Integer
    For i = LBound(ary) To UBound(ary)  'LBound���ŏ��̃C���f�b�N�X�AUBound=�ő�̃C���f�b�N�X
        Debug.Print ary(i)
    Next

End Sub

' Application���x���ō�����
Sub NoDisplay()
    With Application
        .Calculation = xlCalculationManual  '�����v�Z���Ȃ�
        .ScreenUpdating = False '�`�悵�Ȃ�
        .EnableEvents = False   '�Z���̓��e�ύX�ŃC�x���g�𔭐����Ȃ�
    End With
    
    Debug.Print "Do Something"
    
    '�f�t�H���g��Ԃɖ߂�
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub
