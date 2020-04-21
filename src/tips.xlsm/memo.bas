Option Explicit
Sub ボタン左_Click()
    Module1.Import
End Sub

Sub Import()

    ' パス取得
    Dim OpenFileName As String
    OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
    If OpenFileName = "" Then
        Exit Sub
    End If
    
    ' 検証
    If Not Module1.IsValid(OpenFileName) Then
        Exit Sub
        ' MsgBox "Falseの時"
    End If

    ' シートをコピー（上書き）
    Dim destination As Workbook
    Set destination = ThisWorkbook
    
    Dim source As Workbook
    Set source = Workbooks.Open(OpenFileName)
    
    ' あれば消す（消してコピー）
    DeleteIfExists destination, "Data"
    
    ' コピー
    source.Worksheets("Data").Copy after:=destination.Worksheets(destination.Worksheets.Count)
    source.Close
    
    destination.Worksheets(1).Activate

End Sub

Sub DeleteIfExists(bk As Workbook, sheetName As String)

    Dim ws As Worksheet
    Dim exists As Boolean
    For Each ws In Worksheets
        If ws.Name = sheetName Then exists = True
    Next ws

    If exists = True Then
        Application.DisplayAlerts = False
        Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

End Sub

Function IsValid(filename As String) As Boolean
    If Dir(filename) = "" Then
        MsgBox "ファイルが存在しません"
        IsValid = False
        Exit Function
    End If

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = filename Then
            MsgBox filename & vbCrLf & "はすでに開いています", vbExclamation
            IsValid = False
            Exit Function
        End If
    Next wb
    
    IsValid = True
End Function
