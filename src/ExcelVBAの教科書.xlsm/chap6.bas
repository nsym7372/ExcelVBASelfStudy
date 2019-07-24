Attribute VB_Name = "chap6"
Option Explicit

Sub macro1()
Attribute macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    MsgBox "Hello World!!"
End Sub

Sub macro2()
    Application.OnTime Now + TimeValue("00:00:05"), "ShowMsg"
End Sub

Sub ShowMsg()
    MsgBox "Hello World"
End Sub


