Option Explicit

Sub task009_1()
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
        positionCode = getPositionCodeByPositionGroupCode(tbl, i)
        
        For j = LBound(searchValsDeptCode, 1) To UBound(searchValsDeptCode, 1)
            Debug.Print "役職グループ:" & i & "所属コード:" & searchValsDeptCode(j)
            targetPerson = getTargetPerson(tblData, positionCode, searchValsDeptCode(j))
        
            If IsArray(targetPerson) Then
                Debug.Print "実行結果"
                For k = LBound(targetPerson, 2) To UBound(targetPerson, 2)
                    Debug.Print "役職コード:" & targetPerson(12, k) & ":" & targetPerson(1, k) & ":" & targetPerson(2, k)
                Next
            End If
        Next
    Next

End Sub

'役職マスタTBから役職コードに該当する役職を取得する
Function getPositionCodeByPositionGroupCode(tbl As ListObject, searchVal As Long) As Variant
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
    searchCol = tbl.ListColumns("役職グループコード").Index
    targetCol = tbl.ListColumns("役職コード").Index
    
    
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

Function getTargetPerson(tbl As ListObject, searchVals() As Long, searchVals2 As Variant) As Variant
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
    searchCol = tbl.ListColumns("役職コード").Index
    searchCol2 = tbl.ListColumns("所属コード").Index
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

