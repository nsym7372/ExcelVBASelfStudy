VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Goods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private profit_ As Currency
Public name_ As String
Public Price As Currency

'読み取り専用
Public Property Get Profit() As Currency
    Profit = profit_
End Property

'getter
Public Property Get Name() As String
    Name = name_
End Property

'setter
Public Property Let Name(Name As String)
    name_ = Name
End Property



Public Sub ShowInfo()
    MsgBox "商品名：" & name_ & vbCrLf & "価格：" & Price + Profit
End Sub

Public Function ToArray() As Variant
    ToArray = Array(name_, Price)
End Function

' コンストラクタに相当するものはない
' initializeイベントはNew呼ばれた際に実行
Private Sub Class_Initialize()
    Me.name_ = "未設定"
    Me.Price = 99999
    profit_ = 1000
    ' Me.profit_ = 1000　何故かエラー　意味が分からん
End Sub
