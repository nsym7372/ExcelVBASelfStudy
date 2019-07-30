Attribute VB_Name = "ŠÖ”ŒÄ‚Ño‚µ"
Option Explicit

' ”CˆÓ‚Ì”‚Ìˆø”
Sub caller()
    Call callee("‚©‚¦‚é‚Ì", "‚¤‚½‚ª", "‚«‚±‚¦‚Ä", "‚­‚é‚æ")
End Sub

Sub callee(ParamArray values())

    Dim i As Integer
    For i = 0 To UBound(values)
        Debug.Print values(i)
    Next

End Sub

