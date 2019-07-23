Attribute VB_Name = "chap5"
Option Explicit

Sub PrintStrArray(ByRef ary() As String)

    Dim i As Integer
    For i = 0 To UBound(ary)
        Debug.Print ary(i)
    Next

End Sub

'�z��

Sub macro1()
    Dim ary(2) As String

    ary(0) = "�����낤"
    ary(1) = "���낤"
    ary(2) = "���Ԃ낤"
'    ary(3) = "���낤"  �G���[

    Debug.Print "�擪�̃C���f�b�N�X�F" + str(LBound(ary))
    Debug.Print "�����̃C���f�b�N�X�F" + str(UBound(ary))
    
'    ReDim ary(3) ���������ɔz�񐔂����߂Ă��邽�߁A�G���[

    Dim ary2() As String
    ReDim ary2(3) 'ok
    
    ary2(0) = "�����낤"
    ary2(1) = "���낤"
    ary2(2) = "���Ԃ낤"
    ary2(3) = "���낤"
    
    ReDim Preserve ary2(4) '�l�ێ����Ċg��
    ary2(4) = "���낤"
    
    ReDim Preserve ary2(3) '�l�ێ����č폜�����ߕ��͍폜
    
    PrintStrArray ary2
    
End Sub

' split, join
Sub macro2()
    Dim str As String
    str = "�����낤,���낤,���Ԃ낤"
    
    Dim ary() As String
    ary = Split(str, ",")
    
    PrintStrArray ary
    
    Debug.Print Join(ary, ":")
    
End Sub

' �񎟌��z�񁨃Z��
Sub macro3()
    Dim ri As Integer, ci As Integer
    Dim ary(2, 4) As String
    'Dim ary(1 To 3, 1 To 5) As String �������ł�����@�C���f�b�N�X��1����
    
    For ri = 0 To UBound(ary)
        For ci = 0 To UBound(ary, 2)
            ary(ri, ci) = ri & ":" & ci
        Next
    Next
    

    Range("B2").Resize(UBound(ary, 1), UBound(ary, 2)).Value = ary '����߂đ傫���͔z��ɍ��킹��
    'Range("B2:F4") = ary   ����

End Sub

'�Z�����񎟌��z��
Sub macro4()
    Dim ary() As Variant
    ary = Range("B2:F4").Value
    
    Debug.Print ary(1, 1) '�C���f�b�N�X��1����
    
    Dim ary2() As Variant
    ary2 = Range("B2:B5").Value
    Debug.Print UBound(ary2, 1)
    Debug.Print UBound(ary2, 2) '�񎟌��z��ɂȂ��Ă��邽�߁A�G���[�ɂȂ�Ȃ�
    
    Dim ary3() As Variant
    ary3 = Application.WorksheetFunction.Transpose(Range("B2:B5").Value)
    Debug.Print UBound(ary3, 1)
'    Debug.Print UBound(ary3, 2)�@'�z��ׁ̈A�G���[
    
End Sub

'���
Sub macro5()
    Dim ary() As Variant
    ary = Array("�����낤", "���낤", "���Ԃ낤")
    
    Debug.Print ary(0)
    
End Sub

'collection
Sub macro6()
    Dim user As Collection
    Set user = New Collection
    
    user.Add "�����낤"
    user.Add "���낤"
    user.Add "���Ԃ낤"
    
    Debug.Print user.Count
    Debug.Print user(1) '�C���f�b�N�X��1����
    
    user.Add "���낤", after:=1 '�C�ӂ̈ʒu�ɑ}��
    Debug.Print user(2)
End Sub

Sub macro7()
    Dim item As Collection
    Set item = New Collection
    
    ' add( value, key )
    item.Add 100, "�����낤"
    item.Add 200, "���낤"
    item.Add 300, "���Ԃ낤"
    
    Debug.Print item("�����낤")
'    item.Add 400, "�����낤"�@�L�[�d���̓G���[
End Sub

Sub macro8()
     Dim item As Variant
     Set item = CreateObject("Scripting.Dictionary")
     
     item.Add "�����낤", 100
     item.Add "���낤", 200
     item.Add "���Ԃ낤", 300
     
     Debug.Print Join(item.keys, ",")
     Debug.Print Join(item.items, ",")
     
     
End Sub
