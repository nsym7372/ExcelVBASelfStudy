Attribute VB_Name = "データ操作"
Option Explicit

' 配列宣言
Sub DeclareArray()

    '配列宣言には要素数が必要
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"    '当然、型が違うとエラー
    
    Dim s As Variant
    For Each s In ary
        Debug.Print s
    Next
End Sub

' Collection宣言
Sub DeclareCollection()

    'Collectionには要素数不要
    Dim sheets As Collection
    Set sheets = New Collection
    
    With sheets
        .Add Sheet1.Name
        .Add Sheet3.Name
        .Add 1  '型が違えど問題なし
    End With
    
    Dim s As Variant
    For Each s In sheets
        Debug.Print s
    Next
End Sub

' Collectionを連想配列として利用
Sub Hashtable()
    Dim prices As Collection
    Set prices = New Collection
    
    ' 第二引数がキー
    With prices
        .Add 200, "りんご"
        .Add 150, "みかん"
        .Add 500, "ぶどう"
    End With
    
    Debug.Print "りんごの価格：" & prices("りんご")
    
End Sub

' Dictionary
Sub Dictionary()
    Dim prices As Object
    Set prices = CreateObject("Scripting.Dictionary")
    
    With prices
        .Add "りんご", 200
        .Add "みかん", 150
        .Add "ぶどう", 500
    End With
    
    Debug.Print Join(prices.items)
    Debug.Print Join(prices.keys)
    Debug.Print prices.exists("りんご")

End Sub

' For Each
Sub ForEach()
    Dim ary(2) As String
    ary(1) = "Hello"
    ary(2) = "World"
    
    Dim str As Variant 'for eachで使う場合はvariantが必要
    For Each str In ary
        Debug.Print (str)
    Next
End Sub

'要素数で反復
Sub ArrayIndex()
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"
    
    Dim i As Integer
    For i = LBound(ary) To UBound(ary)  'LBound＝最小のインデックス、UBound=最大のインデックス
        Debug.Print ary(i)
    Next

End Sub

' Applicationレベルで高速化
Sub NoDisplay()
    With Application
        .Calculation = xlCalculationManual  '自動計算しない
        .ScreenUpdating = False '描画しない
        .EnableEvents = False   'セルの内容変更でイベントを発生しない
    End With
    
    Debug.Print "Do Something"
    
    'デフォルト状態に戻す
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

' まとめて値を設定
Sub BulkSet()
    ' セルの個数と値の個数が一致していること
    Range("C7:E7").Value = Array(1, 2, 3)
End Sub

' セル結合とClear
Sub MergeAndClear()
    ' 以下二つは同じ結果
    Range("A1").Value = "Hello world"
    Range("A1:B2").Merge
    
    Range("C1").Value = "Hello world"
    Range("C1:D2").Merge

'    Range("A1").ClearContents  結合範囲を全て選択しないとエラー
    Range("A1").MergeArea.ClearContents 'OK
    Range("C1").Value = ""  '見かけ上は同じだが、値は存在する
End Sub


