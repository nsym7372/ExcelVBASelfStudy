Attribute VB_Name = "chap9"
Option Explicit

'macro10 モジュールの先頭に記載が必要
Type Item
    Name As String
    Price As Currency
End Type

Sub macro1()
    Call macro2
    macro2  'call無しでも呼べる
    
    macro3 ("hello")
    
    macro4 ("hello")
    macro4  '引数無い場合は括弧は不可
    
    ' call：基本的には省略可能。省略しない場合は戻り値を受け取れない
    '以下の場合は省略不可であった
    Call macro5(Range("B1:B3"))
    macro5
    Debug.Print "in macro1"
End Sub

Sub macro2()
    Debug.Print "in macro2"
End Sub

Sub macro3(msg As String)
    Debug.Print "in macro3 with " & msg
End Sub

Sub macro4(Optional msg As String = "こんにちはこんにちは")
    Debug.Print "in macro4 with " & msg
End Sub

'Sub macro5(Optional rng As Range = Range("A1:A3")) オブジェクトは代入不可
Sub macro5(Optional rng As Range) '初期値を省略した呼び出しは可能となる
    If rng Is Nothing Then
        Set rng = ActiveCell
    End If
    
    rng.Value = "こんにちは"
    
End Sub

Sub macro6()
    Debug.Print "Hello" + func1() ' 括弧はあってもなくても良い様子
End Sub

'ワークシート数式からも呼べる
Function func1() As String
    func1 = "World!"
End Function

'ワークシート数式からは不可
Private Function func2() As String
    func1 = "World!"
End Function


Sub macro7()
    Dim MyGoods As Goods
    Set MyGoods = New Goods
    
    MyGoods.Name = "フルーツ詰め合わせ"
    MyGoods.Price = 1000
    
    MyGoods.ShowInfo
    Range("A1:B1").Value = MyGoods.ToArray
End Sub

Sub macro8()
    Dim MyGoods As Goods
    Set MyGoods = New Goods
    
    MyGoods.ShowInfo
End Sub

Sub macro9()
    Dim p As IPerson
    Set p = New Person
    
    p.Name = "はまちや"
    p.Say

End Sub

' 構造体の定義はモジュールの先頭
Sub macro10()
    'オブジェクトの定義
    Dim i As Item
    i.Name = "トラックボール"
    i.Price = "5000"
    
    Debug.Print i.Name & ":" & i.Price
End Sub

