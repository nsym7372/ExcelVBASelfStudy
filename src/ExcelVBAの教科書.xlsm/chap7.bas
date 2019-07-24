Attribute VB_Name = "chap7"
Option Explicit

Sub macro4()
    Dim str As String
    str = "World!!"
    
    Debug.Print "Hello"
    
    Stop
    Debug.Print " "
    Debug.Print str
End Sub

Sub macro5()
    Dim str As String
'    str = "World!!"
    
    Debug.Print "Hello"
    Debug.Assert str <> ""
    Debug.Print " "
    Debug.Print str

End Sub

Sub macro6()
    On Error GoTo ErrorHandler
    
    Worksheets("集計").Activate
    Worksheets("集計").Range("A1").Value = 1000
    
    Exit Sub
    
ErrorHandler:

MsgBox "残念なエラー処理"

End Sub

Sub macro7()
    On Error GoTo ErrorHandler
        Worksheets("集計2").Activate
        
    On Error GoTo 0 'resume nextの次に実行される
    
    Worksheets("集計2").Range("A1").Value = 1000
    Exit Sub
    
ErrorHandler:

    Worksheets.Add.Name = "集計2"
    Resume Next
    
End Sub

Sub macro8()
    On Error Resume Next
    Worksheets("集計2").Delete
    
End Sub
