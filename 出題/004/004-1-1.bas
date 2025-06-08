Option Explicit

'内閣府の祝日情報csv
Const URLCSV = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"

Function getHolidaysDay(yearMonth As String) As Object
    Static holidaysInfo As Object
    Dim HolidaysDay As Object: Set HolidaysDay = CreateObject("Scripting.Dictionary")
    Dim targetDateAry As Variant
    Dim csvContent As String
    Dim i As Long
    
    'すべての祝日情報を保持するディクショナリを起動時に生成する
    'static変数で保持しているのでExcel起動中は1回のみ実行される
    If holidaysInfo Is Nothing Then
        csvContent = readCSVFromWeb(URLCSV)
        Set holidaysInfo = getHolidaysDic(csvContent)
    End If
    
    If holidaysInfo.Exists(yearMonth) Then
        targetDateAry = Split(holidaysInfo(yearMonth), ",")
        
        For i = LBound(targetDateAry) To UBound(targetDateAry)
            HolidaysDay.Add Right(targetDateAry(i), 2), ""
        Next i
    Else
        Set HolidaysDay = Nothing
    End If

    Set getHolidaysDay = HolidaysDay

End Function

Function getHolidaysDic(csvContent As String) As Object
    Dim holidaysDic As Object: Set holidaysDic = CreateObject("Scripting.Dictionary")
    Dim lines As Variant
    Dim dateName As Variant

    Dim keyStr As String
    Dim valStr As String
    
    Dim i As Integer
    
    lines = Split(csvContent, vbCrLf)
    
    '先頭はタイトル,末尾は改行が入るのでそれぞれ1ずつ読まない
    For i = 1 To UBound(lines) - 1
        dateName = Split(lines(i), ",")
        keyStr = Format(CDate(dateName(0)), "yyyymm")
        valStr = Format(CDate(dateName(0)), "yyyymmdd")
'        valStr = Format(CDate(dateName(0)), "yyyymmdd") & "-"
'        valStr = valStr + dateName(1)
    
        If holidaysDic.Exists(keyStr) Then
            holidaysDic(keyStr) = holidaysDic(keyStr) & "," & valStr
        Else
            holidaysDic.Add keyStr, valStr
        End If
    Next
    
    Set getHolidaysDic = holidaysDic
End Function

Function readCSVFromWeb(url As String) As String
    Dim httpRequest As Object: Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    
    ' インターネットからCSVをダウンロード
    httpRequest.Open "GET", url, False
    httpRequest.Send
    
    ' ダウンロードした内容を取得
    With stream
       .Type = 1 ' バイナリを読み込む
       .Open
       .Write httpRequest.responseBody
       .Position = 0
       .Type = 2 ' テキストに変更
       .Charset = "shift_jis" '"utf-8"
       readCSVFromWeb = .ReadText
       .Close
    End With
    
    Set stream = Nothing
    Set httpRequest = Nothing
End Function
