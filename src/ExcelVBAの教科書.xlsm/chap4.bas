Attribute VB_Name = "chap4"
Option Explicit

' •¶š—ñ
Const Target As String = "  Hello World!  "

Sub macro1()
    Debug.Print (Len("VBA"))
    Debug.Print (Len("‚±‚ñ‚É‚¿‚Í"))
End Sub

Sub macro2()
    Debug.Print (InStr("hello world!", "o"))
    
    If InStr(Target, "z") = 0 Then
        Debug.Print "z‚ÍŠÜ‚Ü‚ê‚Ä‚¢‚È‚¢"
    End If
End Sub

Sub macro3()

    Debug.Print (Right(Target, 3))
    Debug.Print (Left(Target, 3))
    Debug.Print (Mid(Target, 3, 3))

End Sub

Sub macro4()
    Debug.Print (Trim(Target))
    Debug.Print (RTrim(Target))
    Debug.Print (LTrim(Target))
End Sub

Sub macro5()
    Debug.Print (Replace(Target, "World", "VBA"))
    Debug.Print (StrConv(Target, vbLowerCase))
End Sub

Sub macro6()
    Debug.Print (Format(Date, "ƒtƒ@ƒCƒ‹–¼_yyyymmdd.xl\sx"))
End Sub


' 4-2
'“ú•t
Sub macro11()
    Dim mydate As Date
    mydate = #1/2/2019#
    Debug.Print mydate
End Sub

Sub macro12()
    Debug.Print DateValue("2019”N1Œ2“ú")
    Debug.Print TimeValue("1230•ª50•b")
    
    Debug.Print DateSerial(2019, 1, 2)
    Debug.Print TimeSerial(12, 30, 50)
End Sub

Sub macro13()
    Debug.Print Date
    Debug.Print Month(Date)
    Debug.Print Year(#12/31/2018#)
    Debug.Print Weekday(Date) '1‚ª“ú—j“ú`7‚ª“y—j“ú
End Sub

Sub macro14()
    Dim mydate As Date
    mydate = #1/1/2018#

    Debug.Print DateAdd("yyyy", 1, mydate)  '!= "y"
    Debug.Print DateAdd("m", 1, mydate)
    Debug.Print DateAdd("d", 1, mydate)
End Sub

Sub macro15()
    Dim birthday As Date
    birthday = #1/30/1980#
    Debug.Print DateDiff("yyyy", birthday, Date)
End Sub

Sub macro16()
    Debug.Print Now
    Debug.Print Date
    Debug.Print Time
End Sub
