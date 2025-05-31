Option Explicit

Sub task009_1()
    Dim wsMacro01 As Worksheet
    Dim positionCode() As Long
    Dim targetPerson As Variant
    Dim targetInfo As Variant
    Dim tmpInfo As Variant
    Dim i As Long, j As Long, k As Long, l  As Long
    Dim searchValsDeptCode() As Variant

    searchValsDeptCode = Array(10010, 10020, 10030, 10040, 10050, 10060, 10070, 10080, 10090)

    '役職マスタTB
    Dim tbl As ListObject
    Set tbl = Worksheets("マスタ").ListObjects("役職マスタTB")

    'データTB
    Dim tblData As ListObject
    Set tblData = Worksheets("データ").ListObjects("データTB")

    For i = 1 To 9
        positionCode = getPositionCodeByPositionGroupCode(tbl, i, "役職グループコード", "役職コード")

        For j = LBound(searchValsDeptCode, 1) To UBound(searchValsDeptCode, 1)
            Debug.Print "役職グループ:" & i & "所属コード:" & searchValsDeptCode(j)
            targetPerson = getTargetPerson(tblData, positionCode, "役職コード", searchValsDeptCode(j), "所属コード")

            If IsArray(targetPerson) Then
                'データ用シートを作成していない場合は作成する
                If wsMacro01 Is Nothing Then
                    Set wsMacro01 = createNewSheet
                End If
                
                Call setDataToSheet(wsMacro01, i, CLng(searchValsDeptCode(j)), targetPerson)
                
            End If
        Next
    Next
End Sub

'役職マスタTBから役職コードに該当する役職を取得する
Function getPositionCodeByPositionGroupCode(tbl As ListObject, searchVal As Long, searchHeader As String, targetHeader As String) As Variant
    Dim aryTableData As Variant
    Dim searchCol As Long
    Dim targetCol As Long

    Dim i As Long
    Dim j As Long

    Dim positionCode() As Long
    Dim positionCodeIndex As Long

    ReDim positionCode(8)
    positionCodeIndex = -1

    'テーブル全データを2次元配列取得
    aryTableData = tbl.DataBodyRange.Value
    searchCol = tbl.ListColumns(searchHeader).Index
    targetCol = tbl.ListColumns(targetHeader).Index


    For i = 1 To UBound(aryTableData, 1)
        If aryTableData(i, searchCol) = searchVal Then

           positionCodeIndex = positionCodeIndex + 1

           If positionCodeIndex > UBound(positionCode) Then
               ReDim Preserve positionCode(UBound(positionCode) * 2)
           End If

           positionCode(positionCodeIndex) = aryTableData(i, targetCol)
'           Debug.Print searchVal & vbTab & aryTableData(i, targetCol) & vbTab & aryTableData(i, 2)
        End If
    Next i

    If positionCodeIndex = -1 Then
        getPositionCodeByPositionGroupCode = ""
    Else
        ReDim Preserve positionCode(positionCodeIndex)
        getPositionCodeByPositionGroupCode = positionCode
    End If

End Function


Function getTargetPerson(tbl As ListObject, searchVals() As Long, searchHeader As String, searchVals2 As Variant, searchHeader2 As String) As Variant
    'データテーブル抽出条件
    '役職コード=該当のコード
    Dim aryTableData As Variant

    Dim searchCol As Long
    Dim searchCol2 As Long

    Dim colCount As Long

    Dim aryTableDataCnt As Long
    Dim searchValsCnt As Long
    Dim searchVals2Cnt As Long
    Dim targetPersonCnt As Long

    Dim targetPerson() As Variant
    Dim targetPersonIndex As Long

    Dim isFound As Boolean


    'テーブル全データを2次元配列取得
    aryTableData = tbl.DataBodyRange.Value
    searchCol = tbl.ListColumns(searchHeader).Index
    searchCol2 = tbl.ListColumns(searchHeader2).Index
    colCount = UBound(aryTableData, 2)

    '対象のデータのみを格納するが行列を入れ替えて格納する
    'VBAは最後の次元のみ要素数を変更できる仕様のため
    '列数(ヘッダー)　:固定
    '行数(対象データ数):可変
    '最初は9(2^3+1)個分要素数確保する
    ReDim targetPerson(1 To colCount, 0 To 8)
    targetPersonIndex = -1

    For aryTableDataCnt = LBound(aryTableData, 1) To UBound(aryTableData, 1)
        '抽出条件は同じレベルのインデントにする(抽出条件が増えると深くなりすぎるため)
        'filter conditions are at the same indents level (as it may be too deep when they are too many)
        isFound = False
        '役職コード
        For searchValsCnt = LBound(searchVals) To UBound(searchVals)
        '所属
'        For searchVals2Cnt = LBound(searchVals2) To UBound(searchVals2)
            If aryTableData(aryTableDataCnt, searchCol) = searchVals(searchValsCnt) And _
                aryTableData(aryTableDataCnt, searchCol2) = searchVals2 Then 'searchVals2(searchVals2Cnt) Then

                targetPersonIndex = targetPersonIndex + 1

                If targetPersonIndex > UBound(targetPerson, 2) Then
                    ReDim Preserve targetPerson(1 To colCount, 0 To UBound(targetPerson, 2) * 2)
                End If

                For targetPersonCnt = 1 To colCount
                    '行列を入れ替えて格納する / transpose rows and columns
                    targetPerson(targetPersonCnt, targetPersonIndex) = aryTableData(aryTableDataCnt, targetPersonCnt)
                Next targetPersonCnt

                isFound = True
                Exit For
            End If
'        Next searchVals2Cnt
        If isFound Then Exit For
        Next searchValsCnt
    Next aryTableDataCnt

    If targetPersonIndex = -1 Then
        getTargetPerson = ""
    Else
        ReDim Preserve targetPerson(1 To colCount, 0 To targetPersonIndex)
        getTargetPerson = targetPerson
    End If

End Function


Function createNewSheet() As Worksheet
    Dim ws As Worksheet
    Dim wsName As String
    
    'データ挿入用シートを作成
    wsName = Worksheets("MENU").Range("B4").Value
    'ワークシートを取得してみる
    On Error Resume Next
    Set ws = Worksheets(wsName)
    Select Case Err.Number
        Case 0
            '取得できたとき(シート存在する):シート名+実行日時
            wsName = wsName & "-" & Format(Now(), "yymmdd-hhnnss")
        Case 9
            '変数が空白の時:実行日時
            If wsName = "" Then
                wsName = Format(Now(), "yymmdd-hhnnss")
            '取得できないで変数に入力があるときは、指定のシート名をそのまま使う
            'Else
            End If
    End Select

    Worksheets("マクロ01ひな形").Copy After:=Sheets("マスタ")
    Set ws = ActiveSheet
    ws.Name = wsName

    Set createNewSheet = ws
End Function


Sub setDataToSheet(ws As Worksheet, positionGroupCode As Long, DeptCode As Long, data As Variant)
    Dim rgDataSt As Range
    Dim i As Long, j As Long

    Set rgDataSt = ws.Cells(getRowNumber(positionGroupCode), getColumnNumber(DeptCode))

    j = 0
    For i = LBound(data, 2) To UBound(data, 2)
        rgDataSt.Offset(j, 0).Value = data(2, i)
        j = j + 1
    Next

End Sub


Function getRowNumber(positionGroupCode) As Long
    Select Case positionGroupCode
        Case 1
            getRowNumber = 3
        Case 2
            getRowNumber = 4
        Case 3
            getRowNumber = 7
        Case 4
            getRowNumber = 11
        Case 5
            getRowNumber = 14
        Case 6
            getRowNumber = 16
        Case 7
            getRowNumber = 28
        Case 8
            getRowNumber = 41
        Case 9
            getRowNumber = 52
    End Select
End Function


Function getColumnNumber(DeptCode) As Long
    Select Case DeptCode
        Case 10010
            getColumnNumber = 3
        Case 10020
            getColumnNumber = 4
        Case 10030
            getColumnNumber = 5
        Case 10040
            getColumnNumber = 6
        Case 10050
            getColumnNumber = 7
        Case 10060
            getColumnNumber = 8
        Case 10070
            getColumnNumber = 9
        Case 10080
            getColumnNumber = 10
        Case 10090
            getColumnNumber = 11
    End Select
End Function

