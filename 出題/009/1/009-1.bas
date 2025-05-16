Option Explicit

Sub Sub task009_1()
    Dim positionCode() As Long
    Dim targetPerson() As Variant
    Dim i As Long

    '役職マスタTB
    Dim tbl As ListObject
    Set tbl = Worksheets("マスタ").ListObjects("役職マスタTB")

    'データTB
    Dim tblData As ListObject
    Set tblData = Worksheets("データ").ListObjects("データTB")

    For i = 1 To 9
        positionCode = getPositionCodeByPositionCode(tbl, i, "役職グループコード", "役職コード")
        targetPerson = getTargetPerson(tblData, positionCode, "役職コード")
        
        Debug.Print i & "-->" & UBound(targetPerson, 2) + 1
'        Stop
    Next

End Sub

'役職マスタTBから役職コードに該当する役職を取得する
Function getPositionCodeByPositionCode(tbl As ListObject, searchVal As Long, searchHeader As String, targetHeader As String) As Long()
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
        Erase positionCode
    Else
        ReDim Preserve positionCode(positionCodeIndex)
    End If

    getPositionCodeByPositionCode = positionCode

End Function


Function getTargetPerson(tbl As ListObject, searchVals() As Long, searchHeader As String) As Variant
    'データテーブル抽出条件
    '所属コード=10010, 10020, 10030, 10040, 10050, 10060, 10070, 10080, 10090
    '(役員で表ににない所属コードがあったので実際に必要な所属コードを抽出条件とする)
    '役職コード=該当のコード
    Dim aryTableData As Variant
    Dim searchValsDeptCode As Variant
    Dim searchCol As Long
    Dim searchColFix As Long

    Dim colCount As Long

    Dim aryTableDataCnt As Long
    Dim searchValsCnt As Long
    Dim searchValsDeptCodeCnt As Long
    Dim targetPersonCnt As Long

    Dim targetPerson() As Variant
    Dim targetPersonIndex As Long
    
    Dim isFound As Boolean

    searchValsDeptCode = Array(10010, 10020, 10030, 10040, 10050, 10060, 10070, 10080, 10090)
    'テーブル全データを2次元配列取得
    aryTableData = tbl.DataBodyRange.Value
    searchCol = tbl.ListColumns(searchHeader).Index
    searchColDeptCode = tbl.ListColumns("所属コード").Index
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
        isFound = False
        For searchValsCnt = LBound(searchVals) To UBound(searchVals)
        For searchValsDeptCodeCnt = LBound(searchValsDeptCode) To UBound(searchValsDeptCode)
            If aryTableData(aryTableDataCnt, searchCol) = searchVals(searchValsCnt) And _
                aryTableData(aryTableDataCnt, searchColDeptCode) = searchValsDeptCode(searchValsDeptCodeCnt) Then
                
                targetPersonIndex = targetPersonIndex + 1
    
                If targetPersonIndex > UBound(targetPerson, 2) Then
                    ReDim Preserve targetPerson(1 To colCount, 0 To UBound(targetPerson, 2) * 2)
                End If
    
                For targetPersonCnt = 1 To colCount
                    targetPerson(targetPersonCnt, targetPersonIndex) = aryTableData(aryTableDataCnt, targetPersonCnt)
                Next targetPersonCnt
    
'                Debug.Print searchVals(searchValsCnt) & vbTab & aryTableData(aryTableDataCnt, 1) & vbTab & aryTableData(aryTableDataCnt, 2)
                isFound = True
                Exit For
            End If
        Next searchValsDeptCodeCnt
        If isFound Then Exit For
        Next searchValsCnt
    Next aryTableDataCnt

    If targetPersonIndex = -1 Then
        ReDim targetPerson(0 To -1, 1 To colCount)
    Else
        ReDim Preserve targetPerson(1 To colCount, 0 To targetPersonIndex)
    End If

    getTargetPerson = targetPerson

End Function

