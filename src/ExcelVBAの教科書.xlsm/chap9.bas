Attribute VB_Name = "chap9"
Option Explicit

'macro10 ���W���[���̐擪�ɋL�ڂ��K�v
Type Item
    Name As String
    Price As Currency
End Type

Sub macro1()
    Call macro2
    macro2  'call�����ł��Ăׂ�
    
    macro3 ("hello")
    
    macro4 ("hello")
    macro4  '���������ꍇ�͊��ʂ͕s��
    
    ' call�F��{�I�ɂ͏ȗ��\�B�ȗ����Ȃ��ꍇ�͖߂�l���󂯎��Ȃ�
    '�ȉ��̏ꍇ�͏ȗ��s�ł�����
    Call macro5(Range("B1:B3"))
    macro5
    Debug.Print "in macro1"
End Sub

Sub macro2()
    Debug.Print "in macro2"
End Sub

Sub macro3(msg As String)
    Debug.Print "in macro3 with " & msg
End Sub

Sub macro4(Optional msg As String = "����ɂ��͂���ɂ���")
    Debug.Print "in macro4 with " & msg
End Sub

'Sub macro5(Optional rng As Range = Range("A1:A3")) �I�u�W�F�N�g�͑���s��
Sub macro5(Optional rng As Range) '�����l���ȗ������Ăяo���͉\�ƂȂ�
    If rng Is Nothing Then
        Set rng = ActiveCell
    End If
    
    rng.Value = "����ɂ���"
    
End Sub

Sub macro6()
    Debug.Print "Hello" + func1() ' ���ʂ͂����Ă��Ȃ��Ă��ǂ��l�q
End Sub

'���[�N�V�[�g����������Ăׂ�
Function func1() As String
    func1 = "World!"
End Function

'���[�N�V�[�g��������͕s��
Private Function func2() As String
    func1 = "World!"
End Function


Sub macro7()
    Dim MyGoods As Goods
    Set MyGoods = New Goods
    
    MyGoods.Name = "�t���[�c�l�ߍ��킹"
    MyGoods.Price = 1000
    
    MyGoods.ShowInfo
    Range("A1:B1").Value = MyGoods.ToArray
End Sub

Sub macro8()
    Dim MyGoods As Goods
    Set MyGoods = New Goods
    
    MyGoods.ShowInfo
End Sub

Sub macro9()
    Dim p As IPerson
    Set p = New Person
    
    p.Name = "�͂܂���"
    p.Say

End Sub

' �\���̂̒�`�̓��W���[���̐擪
Sub macro10()
    '�I�u�W�F�N�g�̒�`
    Dim i As Item
    i.Name = "�g���b�N�{�[��"
    i.Price = "5000"
    
    Debug.Print i.Name & ":" & i.Price
End Sub

