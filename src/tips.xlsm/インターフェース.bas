Attribute VB_Name = "インターフェース"
Option Explicit

' インターフェースを使用したポリモーフィズムの実装
Sub InterfaceImplement()
    Dim Cage(1) As IAnimal
   
    Set Cage(0) = New Dog
    Cage(0).subject = "犬"
    
    Set Cage(1) = New Cat
    Cage(1).subject = "猫"
    
    Dim i As Integer
    For i = 0 To UBound(Cage)
        Cage(i).Bark
    Next
    
End Sub



