Attribute VB_Name = "�֐��Ăяo��"
Option Explicit

' �C�ӂ̐��̈���
Sub caller()
    Call callee("�������", "������", "��������", "�����")
End Sub

Sub callee(ParamArray values())

    Dim i As Integer
    For i = 0 To UBound(values)
        Debug.Print values(i)
    Next

End Sub

