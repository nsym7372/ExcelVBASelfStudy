Attribute VB_Name = "chap3"
Option Explicit

'���
Sub macro1()
    ' Dim foo as Long = 10 error!
    
    Dim foo As Long, bar As String
    
    bar = "foo�̒l�F"
    foo = 10
    foo = foo * 5
    Debug.Print bar, foo

    Dim buzz As String: buzz = "���������������Ȃ�OK�����A���@�`�F�b�N���Ȃ����߈Ӗ��Ȃ������E�E�E"
    Debug.Print (buzz)


End Sub


Sub macro2()
    Dim r As Range
    Set r = Range("A1:c3")
    r.Value = "VBA"
End Sub

'�ϐ��錾
Sub macro3()

'   Option Explicit�ŃG���[
'    foo = 10
'    Debug.Print foo
    
'   Option Explicit�ł�����͒ʂ�
    Dim bar
    bar = 20
    Debug.Print bar
    
    Dim buzz As Long
    buzz = 30
    Debug.Print buzz
    
End Sub

'�C���N�������g
Sub macro4()
    Dim i As Integer
    i = 1
    
'    i++    �G���[
'    i += 1 ������G���[
    i = i + 1   '�����OK
    
    Debug.Print i
End Sub

'���
Sub macro5()
    Dim r As Range
    Set r = Range("A1") '�I�u�W�F�N�g�̊i�[
    r = Range("B1") '��xset����ƁA�ȍ~��set�����Ă��ʂ�
    
    Dim r2 As Range
'    r2 = Range("B2") 'set���Ȃ��ƒʂ�Ȃ�
        
    Dim r1 As Object
'    r1 = Range("B3")   �G���[
    Set r1 = Range("B3") 'object�^��Set�K�{
    
    
    Dim i As Integer
    Let i = 10  '���e�����A�ȗ����Ȃ��ꍇ��Let
    
End Sub

'�萔
Sub macro6()
    Const Tax As Double = 0.08  '�萔�͏�����������ł���i���Ȃ���΂Ȃ�Ȃ��j
'    Const Tax as Double    �G���[
    Debug.Print Tax
End Sub

'������A���A�Z�������s
Sub macro7()
    Range("C3").Value = "excel" & "vba"
    Range("C4").Value = "excel" + "vba"
    Range("C5").Value = "excel" + vbLf + "vba"
End Sub

'�I�u�W�F�N�g��r
Sub macro8()
    MsgBox Sheets(1) Is Worksheets("sheet1")
    
    If Cells.Find("C#") Is Nothing Then
        MsgBox "���݂��Ȃ�"
    End If
        
    ' �s�xrange�I�u�W�F�N�g���擾���邽�߁A�ʃI�u�W�F�N�g�ƂȂ�
    MsgBox Range("A1") Is Range("A1")
    
    ' ����Z�����̔�r�́A�Ԓn���r
    MsgBox Range("A1").Address = Range("A1").Address
    
End Sub

' �_�����Z�q
Sub macro9()
    ' �_���ς�And�A�_���a��Or
    ' �ے艉�Z�q��Not
    If Not Cells.Find("VBA") Is Nothing Then
        MsgBox "exists!"
    End If

End Sub

' foreach
Sub macro10()
    Dim r As Range
    For Each r In Range("A1:A10")
        r.Value = 1
    Next
End Sub

' �O�����Z�q
Sub macro11()
    MsgBox IIf(Cells.Find("VBA") Is Nothing, "�Ȃ�����", "������")

End Sub

' select�i�������swich)
Sub macro12()

    Select Case Range("A1").Value
        Case "VBA"
            MsgBox "����������"
        Case "C#"
            MsgBox "�^�C�v�Z�[�t"
        Case "Python"
            MsgBox "�l�H�m�\"
            
        Case "VBA" '�ǂ����ň�v����΁A�t�H�[���X���[���Ȃ��i�����͒ʂ�Ȃ��j
            MsgBox "�Â���"
        Case Else
            MsgBox "���̑�"
    End Select
        
End Sub

'confirm
Sub macro13()
    
    Dim ret As VbMsgBoxResult
    ret = MsgBox("�ŋ߂ǂ��H", Buttons:=vbYesNo, Title:="�܂���")
    
    If ret = vbYes Then
        MsgBox "���\�ł��Ȃ�"
    Else
        MsgBox "����΂��["
    End If
    
    
End Sub

' �l�̓���
Sub macro14()
    Dim ret As String
    ret = InputBox("����ɂ��͂���ɂ���")
    
    MsgBox "�ڂ�" + ret + "�����!"
End Sub

' �Z���͈͂�I��������
Sub macro15()
    Dim selected As Range, r As Range
    Set selected = Application.InputBox("�Ώ۔͈͂�I��", Type:=8)
    
    For Each r In selected
        r.Value = "�Z"
    Next
    
    
End Sub
