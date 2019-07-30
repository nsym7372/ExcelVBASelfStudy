Attribute VB_Name = "Œ©‰h‚¦"
Option Explicit

' •‚ğ©“®’²®‚µ‚ÄA‚¿‚å‚Á‚Æ—]”’‚ğ‰Á‚¦‚é
Sub ColWidth()
    
    Dim rng As Range
    With Range("C:E")
        .EntireColumn.AutoFit
        For Each rng In .Columns
            rng.ColumnWidth = rng.ColumnWidth + 2
        Next
    End With

End Sub
