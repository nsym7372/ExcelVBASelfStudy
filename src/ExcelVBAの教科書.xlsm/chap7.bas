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
    
    Worksheets("�W�v").Activate
    Worksheets("�W�v").Range("A1").Value = 1000
    
    Exit Sub
    
ErrorHandler:

MsgBox "�c�O�ȃG���[����"

End Sub

Sub macro7()
    On Error GoTo ErrorHandler
        Worksheets("�W�v2").Activate
        
    On Error GoTo 0 'resume next�̎��Ɏ��s�����
    
    Worksheets("�W�v2").Range("A1").Value = 1000
    Exit Sub
    
ErrorHandler:

    Worksheets.Add.Name = "�W�v2"
    Resume Next
    
End Sub

Sub macro8()
    On Error Resume Next
    Worksheets("�W�v2").Delete
    
End Sub
