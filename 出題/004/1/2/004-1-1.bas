Option Explicit

'内閣府の祝日情報csv
Const URLCSV = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
Const URLICS = "https://calendar.google.com/calendar/ical/ja.japanese%23holiday%40group.v.calendar.google.com/public/basic.ics"

Const wsName = "祝日"

Const StartWithDateLine As String = "DTSTART;VALUE=DATE:"
Const StartWithNameLine As String = "SUMMARY:"

Enum ColHolidays
    holidayDate = 1
    holidayName
    holidayStart = 1
    holidayEnd = 2
End Enum

Sub setHolidaysInfoToSh()
    Worksheets(wsName).Range(Worksheets(wsName).Columns(ColHolidays.holidayStart), Worksheets(wsName).Columns(ColHolidays.holidayEnd)).Clear

    Call setNaikakufuHolidaysInfoToSh
    Call setGoogleHolidaysInfoToSh

    With Worksheets(wsName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Worksheets(wsName).Columns(ColHolidays.holidayDate), Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Worksheets(wsName).Range(Worksheets(wsName).Columns(ColHolidays.holidayStart), Worksheets(wsName).Columns(ColHolidays.holidayEnd))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Worksheets(wsName).Range("D1").Value = Format(Now(), "yyyy/mm/dd")
End Sub

Sub setNaikakufuHolidaysInfoToSh()
    Dim contentCsv As String
    
    Dim lines As Variant
    Dim dateName As Variant
    
    Dim i As Integer
    
    
    contentCsv = readTextFromWeb(URLCSV, "shift_jis")
    
    lines = Split(contentCsv, vbCrLf)
    
    '先頭はタイトル,末尾は改行が入るのでそれぞれ1ずつ読まない
    For i = 1 To UBound(lines) - 1
        dateName = Split(lines(i), ",")
        With Worksheets(wsName)
            .Cells(i, ColHolidays.holidayDate) = dateName(0)
            .Cells(i, ColHolidays.holidayName) = dateName(1)
        End With
    Next

End Sub

'Googleカレンダーでは法律上の祝日でない祝日が含まれるため(節分,雛祭りなど)
'内閣府csvに登録された最終年から祝日名を取得して同じ祝日名(と振替休日)のみ法定祝日と判定する
Sub setGoogleHolidaysInfoToSh()
    Dim lastRow As Long
    Dim lastYear As Long
    Dim targetHolidayNames As String
    
    Dim contentIcs As String
    
    Dim lines As Variant
    Dim ymdStr As String
    
    Dim i As Long, j As Long
    
    lastRow = Worksheets(wsName).Cells(rows.Count, 1).End(xlUp).Row
    lastYear = Year(Worksheets(wsName).Cells(lastRow, 1).Value)
    
    i = lastRow
    Do While Year(Worksheets(wsName).Cells(i, 1).Value) = lastYear
        targetHolidayNames = targetHolidayNames & Worksheets(wsName).Cells(i, 2).Value & ","
        i = i - 1
    Loop
    
    contentIcs = readTextFromWeb(URLICS, "utf-8")
    
    lines = Split(contentIcs, vbCrLf)
    
    j = lastRow + 1 'エクセル行用のカウンタ変数
    For i = 0 To UBound(lines)
        With Worksheets(wsName)
            If (Left(lines(i), Len(StartWithDateLine)) = StartWithDateLine) Then
                ymdStr = Mid(lines(i), Len(StartWithDateLine) + 1)
                If CLng(Left(ymdStr, 4)) > lastYear Then
                    .Cells(j, ColHolidays.holidayDate) = DateSerial(Left(ymdStr, 4), Mid(ymdStr, 5, 2), Right(ymdStr, 2))
                End If
            ElseIf (Left(lines(i), Len(StartWithNameLine))) = StartWithNameLine Then
                If .Cells(j, ColHolidays.holidayDate) <> "" Then
                    '前年と同じ祝日名/振替休日か確認する
                    If InStr(targetHolidayNames, Mid(lines(i), Len(StartWithNameLine) + 1, Len(lines(i)))) > 0 Or _
                       InStr(Mid(lines(i), Len(StartWithNameLine) + 1, Len(lines(i))), "休日") > 0 Then
                        .Cells(j, ColHolidays.holidayName) = Mid(lines(i), Len(StartWithNameLine) + 1, Len(lines(i)))
                        j = j + 1
                    Else
                        Debug.Print Mid(lines(i), Len(StartWithNameLine) + 1, Len(lines(i)))
                        .Cells(j, ColHolidays.holidayDate) = ""
                    End If
                End If
            End If
        End With
    Next
End Sub

Function readTextFromWeb(url As String, charCode As String) As String
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
       .Charset = charCode
       readTextFromWeb = .ReadText
       .Close
    End With
    
    Set stream = Nothing
    Set httpRequest = Nothing
End Function

Function getHolidaysDay(yearMonth As String) As Object
    Static holidaysInfo As Object
    Dim HolidaysDay As Object: Set HolidaysDay = CreateObject("Scripting.Dictionary")
    Dim targetDateAry As Variant
    Dim csvContent As String
    Dim i As Long

    'すべての祝日情報を保持するディクショナリを起動時に生成する
    'static変数で保持しているのでExcel起動中は1回のみ実行される
    If holidaysInfo Is Nothing Then
        Set holidaysInfo = getHolidaysDic()
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

Function getHolidaysDic() As Object
    Dim holidaysDic As Object: Set holidaysDic = CreateObject("Scripting.Dictionary")

    Dim keyStr As String
    Dim valStr As String

    Dim i As Integer
    
    '最終更新日より1年以上経っていたら祝日情報を更新する
    If Worksheets(wsName).Range("D1").Value < DateAdd("yyyy", -1, Date) Then
        Call setHolidaysInfoToSh
    End If
    
    Worksheets(wsName).Range("D1").Value = Format(Now(), "yyyy/mm/dd")

    i = 1
    Do Until Worksheets(wsName).Cells(i, 1).Value = ""
        keyStr = Format(Worksheets(wsName).Cells(i, 1).Value, "yyyymm")
        valStr = Format(Worksheets(wsName).Cells(i, 1).Value, "dd")

        If holidaysDic.Exists(keyStr) Then
            holidaysDic(keyStr) = holidaysDic(keyStr) & "," & valStr
        Else
            holidaysDic.Add keyStr, valStr
        End If
        i = i + 1
    Loop

    Set getHolidaysDic = holidaysDic
End Function
