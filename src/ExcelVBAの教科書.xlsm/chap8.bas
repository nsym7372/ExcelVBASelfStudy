Attribute VB_Name = "chap8"
Option Explicit

Sub macro1()
    Dim spvoice As Object
    Set spvoice = CreateObject("SAPI.SpVoice")
    spvoice.Speak "hello excel!!"
End Sub

Sub macro2()
    
    Dim ws As Object
    
    '�Q�Ɛݒ肵����́ANew���Z�q��OK
    'Set ws = CreateObject("WScript.Network")
    Set ws = New IWshRuntimeLibrary.WshNetwork
    
    Debug.Print ws.ComputerName
End Sub
