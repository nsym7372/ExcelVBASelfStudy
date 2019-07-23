Attribute VB_Name = "chap2"
Sub macro1()
   Debug.Print "hello vba!"
End Sub

Sub macro2()
    For i = 1 To 10
        Cells(i, 1).Value = i * 10
    Next
End Sub

Sub macro3()
    Range("A4").Value = "hello vba!"
    Cells(3, 1).Value = 1000
    Range("A1:B2").Value = #6/5/2018#
    
    Range("B2").ClearContents
End Sub

Sub macro4()
    Debug.Print Range("A1").Width
End Sub

Sub macro5()
    Range("A5").Interior.Color = RGB(255, 0, 0)
    Range("A5").Value = "color red"
    Range("A5").ClearContents
End Sub

Sub macro6()
    Range("E1:G3").Value = Array(1, 2, 3)
    Range("F2").Delete
    
    
    Range("I1:K3").Value = Array(1, 2, 3)
    Range("J2").Delete shift:=xlShiftToLeft
    
End Sub



