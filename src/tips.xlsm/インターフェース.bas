Attribute VB_Name = "�C���^�[�t�F�[�X"
Option Explicit

' �C���^�[�t�F�[�X���g�p�����|�����[�t�B�Y���̎���
Sub InterfaceImplement()
    Dim Cage(1) As IAnimal
   
    Set Cage(0) = New Dog
    Cage(0).subject = "��"
    
    Set Cage(1) = New Cat
    Cage(1).subject = "�L"
    
    Dim i As Integer
    For i = 0 To UBound(Cage)
        Cage(i).Bark
    Next
    
End Sub



