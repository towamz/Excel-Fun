'yyyymm形式のテキストか確認する
Function isYearMonthStr(yearMonthStr As String) As Boolean
    Dim Dt As Date
    
    If Len(yearMonthStr) <> 6 Then
        isYearMonthStr = False
        Exit Function
    End If
    
    'dateSerialでエラーが発生しないか確認
    On Error Resume Next
    Dt = DateSerial(Left(yearMonthStr, 4), Right(yearMonthStr, 2), 1)
    If Err.Number <> 0 Then
        isYearMonthStr = False
        Exit Function
    End If
    On Error GoTo 0

    '通常の年月でないyyyymmを指定していると一致しないためチェック
    '例: 202513 -> 202601になる
    If yearMonthStr <> Format(Dt, "yyyymm") Then
        isYearMonthStr = False
        Exit Function
    End If

    isYearMonthStr = True
End Function

'翌年月・前年月を取得する
Function getOffsetYearMonthStr(yearMonthStr As String, yearOffset As Long, monthOffset As Long) As String
    '引数がyyyymm形式であれば処理実行
    If isYearMonthStr(yearMonthStr) Then
        getOffsetYearMonthStr = Format(DateSerial(Left(yearMonthStr, 4) + yearOffset, Right(yearMonthStr, 2) + monthOffset, 1), "yyyymm")
    '引数がyyyymm形式以外であれば引数をそのまま返す
    Else
        getOffsetYearMonthStr = yearMonthStr
    End If
End Function
