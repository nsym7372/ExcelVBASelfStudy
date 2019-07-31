Attribute VB_Name = "�f�[�^����"
Option Explicit

' �z��錾
Sub DeclareArray()

    '�z��錾�ɂ͗v�f�����K�v
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"    '���R�A�^���Ⴄ�ƃG���[
    
    Dim s As Variant
    For Each s In ary
        Debug.Print s
    Next
End Sub

' Collection�錾
Sub DeclareCollection()

    'Collection�ɂ͗v�f���s�v
    Dim sheets As Collection
    Set sheets = New Collection
    
    With sheets
        .Add Sheet1.Name
        .Add Sheet3.Name
        .Add 1  '�^���Ⴆ�ǖ��Ȃ�
    End With
    
    Dim s As Variant
    For Each s In sheets
        Debug.Print s
    Next
End Sub

' Collection��A�z�z��Ƃ��ė��p
Sub Hashtable()
    Dim prices As Collection
    Set prices = New Collection
    
    ' ���������L�[
    With prices
        .Add 200, "���"
        .Add 150, "�݂���"
        .Add 500, "�Ԃǂ�"
    End With
    
    Debug.Print "��񂲂̉��i�F" & prices("���")
    
End Sub

' Dictionary
Sub Dictionary()
    Dim prices As Object
    Set prices = CreateObject("Scripting.Dictionary")
    
    With prices
        .Add "���", 200
        .Add "�݂���", 150
        .Add "�Ԃǂ�", 500
    End With
    
    Debug.Print Join(prices.items)
    Debug.Print Join(prices.keys)
    Debug.Print prices.exists("���")

End Sub

' For Each
Sub ForEach()
    Dim ary(2) As String
    ary(1) = "Hello"
    ary(2) = "World"
    
    Dim str As Variant 'for each�Ŏg���ꍇ��variant���K�v
    For Each str In ary
        Debug.Print (str)
    Next
End Sub

'�v�f���Ŕ���
Sub ArrayIndex()
    Dim ary(3) As String
    ary(1) = Sheet1.Name
    ary(2) = Sheet3.Name
    ary(3) = "1"
    
    Dim i As Integer
    For i = LBound(ary) To UBound(ary)  'LBound���ŏ��̃C���f�b�N�X�AUBound=�ő�̃C���f�b�N�X
        Debug.Print ary(i)
    Next

End Sub

' Application���x���ō�����
Sub NoDisplay()
    With Application
        .Calculation = xlCalculationManual  '�����v�Z���Ȃ�
        .ScreenUpdating = False '�`�悵�Ȃ�
        .EnableEvents = False   '�Z���̓��e�ύX�ŃC�x���g�𔭐����Ȃ�
    End With
    
    Debug.Print "Do Something"
    
    '�f�t�H���g��Ԃɖ߂�
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

' �܂Ƃ߂Ēl��ݒ�
Sub BulkSet()
    ' �Z���̌��ƒl�̌�����v���Ă��邱��
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
        If matches.Count Then ' success����
            Debug.Print matches(0) ' �S��
            Debug.Print matches(0).SubMatches(0) ' �L���v�`���������̂��擾
        End If
    End With
End Sub

' �l�̃R�s�[
Sub CopyData()
     
    With ThisWorkbook.Worksheets("Sheet1")
        .Range("A1:A10").Copy (.Range("C3"))
    End With
    
End Sub

' �V�[�g����āA������V�[�g�ɒl���R�s�[
Sub CreateBookAndAddSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

'   �V�����u�b�N�ɃV�[�g��ǉ����āA�l���R�s�[
    With Workbooks.Add
        ws.Range("A1:A10").Copy .Worksheets(1).Range("A1")
        .SaveAs (ThisWorkbook.Path & "\book1.xlsx")
    End With
        
End Sub

