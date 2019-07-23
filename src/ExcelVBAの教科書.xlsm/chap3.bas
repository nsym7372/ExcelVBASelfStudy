Attribute VB_Name = "chap3"
Option Explicit

'代入
Sub macro1()
    ' Dim foo as Long = 10 error!
    
    Dim foo As Long, bar As String
    
    bar = "fooの値："
    foo = 10
    foo = foo * 5
    Debug.Print bar, foo

    Dim buzz As String: buzz = "こういう書き方ならOKだが、文法チェックがないため意味ないかも・・・"
    Debug.Print (buzz)


End Sub


Sub macro2()
    Dim r As Range
    Set r = Range("A1:c3")
    r.Value = "VBA"
End Sub

'変数宣言
Sub macro3()

'   Option Explicitでエラー
'    foo = 10
'    Debug.Print foo
    
'   Option Explicitでもこれは通る
    Dim bar
    bar = 20
    Debug.Print bar
    
    Dim buzz As Long
    buzz = 30
    Debug.Print buzz
    
End Sub

'インクリメント
Sub macro4()
    Dim i As Integer
    i = 1
    
'    i++    エラー
'    i += 1 これもエラー
    i = i + 1   'これはOK
    
    Debug.Print i
End Sub

'代入
Sub macro5()
    Dim r As Range
    Set r = Range("A1") 'オブジェクトの格納
    r = Range("B1") '一度setすると、以降はset無くても通る
    
    Dim r2 As Range
'    r2 = Range("B2") 'setしないと通らない
        
    Dim r1 As Object
'    r1 = Range("B3")   エラー
    Set r1 = Range("B3") 'object型もSet必須
    
    
    Dim i As Integer
    Let i = 10  'リテラル、省略しない場合はLet
    
End Sub

'定数
Sub macro6()
    Const Tax As Double = 0.08  '定数は初期化時代入できる（しなければならない）
'    Const Tax as Double    エラー
    Debug.Print Tax
End Sub

'文字列連結、セル内改行
Sub macro7()
    Range("C3").Value = "excel" & "vba"
    Range("C4").Value = "excel" + "vba"
    Range("C5").Value = "excel" + vbLf + "vba"
End Sub

'オブジェクト比較
Sub macro8()
    MsgBox Sheets(1) Is Worksheets("sheet1")
    
    If Cells.Find("C#") Is Nothing Then
        MsgBox "存在しない"
    End If
        
    ' 都度rangeオブジェクトを取得するため、別オブジェクトとなる
    MsgBox Range("A1") Is Range("A1")
    
    ' 同一セルかの比較は、番地を比較
    MsgBox Range("A1").Address = Range("A1").Address
    
End Sub

' 論理演算子
Sub macro9()
    ' 論理積はAnd、論理和はOr
    ' 否定演算子はNot
    If Not Cells.Find("VBA") Is Nothing Then
        MsgBox "exists!"
    End If

End Sub

' foreach
Sub macro10()
    Dim r As Range
    For Each r In Range("A1:A10")
        r.Value = 1
    Next
End Sub

' 三項演算子
Sub macro11()
    MsgBox IIf(Cells.Find("VBA") Is Nothing, "なかった", "あった")

End Sub

' select（他言語のswich)
Sub macro12()

    Select Case Range("A1").Value
        Case "VBA"
            MsgBox "もうええわ"
        Case "C#"
            MsgBox "タイプセーフ"
        Case "Python"
            MsgBox "人工知能"
            
        Case "VBA" 'どこかで一致すれば、フォールスルーしない（ここは通らない）
            MsgBox "古いわ"
        Case Else
            MsgBox "その他"
    End Select
        
End Sub

'confirm
Sub macro13()
    
    Dim ret As VbMsgBoxResult
    ret = MsgBox("最近どう？", Buttons:=vbYesNo, Title:="まいど")
    
    If ret = vbYes Then
        MsgBox "結構ですなぁ"
    Else
        MsgBox "がんばりやー"
    End If
    
    
End Sub

' 値の入力
Sub macro14()
    Dim ret As String
    ret = InputBox("こんにちはこんにちは")
    
    MsgBox "ぼく" + ret + "ちゃん!"
End Sub

' セル範囲を選択させる
Sub macro15()
    Dim selected As Range, r As Range
    Set selected = Application.InputBox("対象範囲を選択", Type:=8)
    
    For Each r In selected
        r.Value = "〇"
    Next
    
    
End Sub
