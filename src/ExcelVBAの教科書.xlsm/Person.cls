VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerson
'Public Name As String
'Public Function Say()

Private name_ As String '独自メンバ変数

' overrideは IPperson_(メンバ名)
Private Property Let IPerson_Name(ByVal pName As String)
    name_ = pName
End Property

Private Property Get IPerson_Name() As String
    IPerson_Name = name_
End Property

Public Function IPerson_Say()
    Debug.Print name_, " say 「こんにちはこんにちは!」"
End Function


