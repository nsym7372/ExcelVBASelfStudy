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

' Regex
Sub RegularExpression()

    Dim regex As Object
    Dim matches As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = True
        .Pattern = "(\w+)@\w+"
    
        Set matches = .Execute("sample_test@contoso.com")
        If matches.Count Then ' success判定
            Debug.Print matches(0) ' 全体
            Debug.Print matches(0).SubMatches(0) ' キャプチャしたものを取得
        End If
    End With
End Sub

' 値のコピー
Sub CopyData()
     
    With ThisWorkbook.Worksheets("Sheet1")
        .Range("A1:A10").Copy (.Range("C3"))
    End With
    
End Sub

' シート作って、作ったシートに値をコピー
Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   新しいブックにシートを追加して、値をコピー
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
End Sub

