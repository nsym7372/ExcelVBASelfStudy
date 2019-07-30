VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAnimal

Private subject_ As String

Public Property Let IAnimal_subject(ByVal subject As String)
    subject_ = subject
End Property

Public Property Get IAnimal_subject() As String
    IAnimal_subject = subject_
End Property

Public Function IAnimal_Bark()
    MsgBox subject_ & "Ç…Ç·Å["
End Function
