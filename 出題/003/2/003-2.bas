Sub task003_02()
    Dim targetRange As Range
    Dim targetInfo(4) As Variant
    Dim i As Long
    
    '5行目から開始
    i = 5
    Do Until Worksheets("販売データ").Cells(i, 2) = ""
        DoEvents
        '一括出力=1
        If Worksheets("販売データ").Cells(i, 11).Value = 1 Then
            '出力日に入力がない(処理実行)
            If Worksheets("販売データ").Cells(i, 10).Value = "" Then
                Set targetRange = Worksheets("販売データ").Range(Worksheets("販売データ").Cells(i, 2), Worksheets("販売データ").Cells(i, 10))
                'PDF作成に必要なデータを配列に格納する
                targetInfo(0) = targetRange(1, 1)
                targetInfo(1) = targetRange(1, 2)
                targetInfo(2) = getTargetDate(Worksheets("販売データ").Range("F2").Value, targetRange(1, 3), targetRange(1, 4))
                targetInfo(3) = targetRange(1, 8)
                targetInfo(4) = targetRange(1, 5)
        
                If makeBill(targetInfo) Then
                    Worksheets("販売データ").Cells(i, 10).Value = Format(Now(), "m/d") & "済"
                    Worksheets("販売データ").Cells(i, 11).Value = "完了"
                Else
                    Worksheets("販売データ").Cells(i, 11).Value = "失敗"
                End If
            '出力日に入力がある場合(処理実行しない)
            Else
                Worksheets("販売データ").Cells(i, 11).Value = "失敗"
            End If
        '1以外(完了・失敗)
        Else
            '前回の実行結果を消去する
            Worksheets("販売データ").Cells(i, 11).Value = ""
        End If
    
        i = i + 1
    Loop

    Worksheets("販売データ").Activate
    MsgBox "PDF出力しました"

End Sub


Sub clearFlg()
     
    Worksheets("販売データ").Range(Worksheets("販売データ").Cells(5, 11), Worksheets("販売データ").Cells(Rows.Count, 11)).ClearContents

    MsgBox "一括フラグをクリアしました", vbOKOnly + vbInformation

End Sub


Function getTargetDate(fiscalYear As Long, month As Long, day As Long) As Date
    Dim year As Long
    
    '1~3月の時は+1する
    If month <= 3 Then
        year = fiscalYear + 1
    Else
        year = fiscalYear
    End If

    getTargetDate = DateSerial(year, month, day)

End Function


Function makeBill(targetInfo As Variant) As Boolean
    Dim PDF As New clsSavePdf
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo err
    'データを挿入する
    Worksheets("領収書").Range("D3") = targetInfo(0)
    Worksheets("領収書").Range("C7") = targetInfo(1)
    Worksheets("領収書").Range("F3") = Format(targetInfo(2), "ggge年mm月dd日")
    Worksheets("領収書").Range("F9") = targetInfo(3)
    Worksheets("領収書").Range("F11") = targetInfo(4)

    'PDFを作成する
    PDF.TargetDirectory = FSO.BuildPath(ThisWorkbook.Path, "領収書")
    PDF.WsName = "領収書"
    PDF.PdfName = "領収書" & Format(targetInfo(2), "ggge年mm月dd日") & "(" & targetInfo(1) & ")"
    PDF.savePDF

    makeBill = True
    Exit Function

err:
    makeBill = False

End Function

