Attribute VB_Name = "chap5"
Option Explicit

Sub PrintStrArray(ByRef ary() As String)

    Dim i As Integer
    For i = 0 To UBound(ary)
        Debug.Print ary(i)
    Next

End Sub

'配列

Sub macro1()
    Dim ary(2) As String

    ary(0) = "いちろう"
    ary(1) = "じろう"
    ary(2) = "さぶろう"
'    ary(3) = "しろう"  エラー

    Debug.Print "先頭のインデックス：" + str(LBound(ary))
    Debug.Print "末尾のインデックス：" + str(UBound(ary))
    
'    ReDim ary(3) 初期化時に配列数を決めているため、エラー

    Dim ary2() As String
    ReDim ary2(3) 'ok
    
    ary2(0) = "いちろう"
    ary2(1) = "じろう"
    ary2(2) = "さぶろう"
    ary2(3) = "しろう"
    
    ReDim Preserve ary2(4) '値保持して拡張
    ary2(4) = "ごろう"
    
    ReDim Preserve ary2(3) '値保持して削除→超過分は削除
    
    PrintStrArray ary2
    
End Sub

' split, join
Sub macro2()
    Dim str As String
    str = "いちろう,じろう,さぶろう"
    
    Dim ary() As String
    ary = Split(str, ",")
    
    PrintStrArray ary
    
    Debug.Print Join(ary, ":")
    
End Sub

' 二次元配列→セル
Sub macro3()
    Dim ri As Integer, ci As Integer
    Dim ary(2, 4) As String
    'Dim ary(1 To 3, 1 To 5) As String こっちでも動作　インデックスは1から
    
    For ri = 0 To UBound(ary)
        For ci = 0 To UBound(ary, 2)
            ary(ri, ci) = ri & ":" & ci
        Next
    Next
    

    Range("B2").Resize(UBound(ary, 1), UBound(ary, 2)).Value = ary '基準決めて大きさは配列に合わせる
    'Range("B2:F4") = ary   同じ

End Sub

'セル→二次元配列
Sub macro4()
    Dim ary() As Variant
    ary = Range("B2:F4").Value
    
    Debug.Print ary(1, 1) 'インデックスは1から
    
    Dim ary2() As Variant
    ary2 = Range("B2:B5").Value
    Debug.Print UBound(ary2, 1)
    Debug.Print UBound(ary2, 2) '二次元配列になっているため、エラーにならない
    
    Dim ary3() As Variant
    ary3 = Application.WorksheetFunction.Transpose(Range("B2:B5").Value)
    Debug.Print UBound(ary3, 1)
'    Debug.Print UBound(ary3, 2)　'配列の為、エラー
    
End Sub

'代入
Sub macro5()
    Dim ary() As Variant
    ary = Array("いちろう", "じろう", "さぶろう")
    
    Debug.Print ary(0)
    
End Sub

'collection
Sub macro6()
    Dim user As Collection
    Set user = New Collection
    
    user.Add "いちろう"
    user.Add "じろう"
    user.Add "さぶろう"
    
    Debug.Print user.Count
    Debug.Print user(1) 'インデックスは1から
    
    user.Add "しろう", after:=1 '任意の位置に挿入
    Debug.Print user(2)
End Sub

Sub macro7()
    Dim item As Collection
    Set item = New Collection
    
    ' add( value, key )
    item.Add 100, "いちろう"
    item.Add 200, "じろう"
    item.Add 300, "さぶろう"
    
    Debug.Print item("いちろう")
'    item.Add 400, "いちろう"　キー重複はエラー
End Sub

Sub macro8()
     Dim item As Variant
     Set item = CreateObject("Scripting.Dictionary")
     
     item.Add "いちろう", 100
     item.Add "じろう", 200
     item.Add "さぶろう", 300
     
     Debug.Print Join(item.keys, ",")
     Debug.Print Join(item.items, ",")
     
     
End Sub
