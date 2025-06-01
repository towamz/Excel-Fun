Sub switchButton()
    Dim btn As Object

    Set btn = Worksheets("実行").Shapes("Button1")

    Select Case btn.TextFrame.Characters.Text
        Case "未実行"
            btn.TextFrame.Characters.Text = "実行中"
        Case "実行中"
            btn.TextFrame.Characters.Text = "未実行"
        Case Else
            btn.TextFrame.Characters.Text = "未実行"
    End Select

End Sub


Sub task003_01(target As Range)
    Dim btn As Object
    Dim targetRange As Range
    Dim targetInfo(4) As Variant
    Dim msg As String

    Set btn = Worksheets("実行").Shapes("Button1")
    
    If Application.Intersect(Range("B5:B1000"), target) Is Nothing Then
        Exit Sub
    End If

    If btn.TextFrame.Characters.Text <> "実行中" Then
        Debug.Print "実行中に切り替えてください"
        Exit Sub
    End If
    
    If Worksheets("販売データ").Cells(target.Row, 10).Value <> "" Then
        Exit Sub
    End If

    Set targetRange = Worksheets("販売データ").Range(Cells(target.Row, 2), Cells(target.Row, 10))

    'PDF作成に必要なデータを配列に格納する
    targetInfo(0) = targetRange(1, 1)
    targetInfo(1) = targetRange(1, 2)
    targetInfo(2) = getTargetDate(Worksheets("販売データ").Range("F2").Value, targetRange(1, 3), targetRange(1, 4))
    targetInfo(3) = targetRange(1, 8)
    targetInfo(4) = targetRange(1, 5)
    
    msg = msg & "    No:" & targetInfo(0) & vbCrLf
    msg = msg & "購入者:" & targetInfo(1) & vbCrLf
    msg = msg & "年月日:" & Format(targetInfo(2), "ggge年mm月dd日") & vbCrLf
    msg = msg & "  売上:" & targetInfo(3) & vbCrLf
    msg = msg & "  品物:" & targetInfo(4) & vbCrLf

    If MsgBox(msg, vbYesNo + vbQuestion, "下記の領収書PDFを作成しますか") = vbNo Then
        Exit Sub
    End If
    
    
    If makeBill(targetInfo) Then
        Worksheets("販売データ").Activate
        Worksheets("販売データ").Cells(target.Row, 10).Value = Format(Now(), "m/d") & "済"
        MsgBox "PDF出力しました"
    End If

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

