Attribute VB_Name = "関数呼び出し"
Option Explicit

' 任意の数の引数
Sub caller()
    Call callee("かえるの", "うたが", "きこえて", "くるよ")
End Sub

Sub callee(ParamArray values())

    Dim i As Integer
    For i = 0 To UBound(values)
        Debug.Print values(i)
    Next

End Sub

