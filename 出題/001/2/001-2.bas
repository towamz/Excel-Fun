Option Explicit

Sub task001_2()
    Dim ws As Worksheet
    Dim rg As Range

    Set ws = Worksheets("スケジュール")
    Set rg = ws.Range("B3")

    Call setColor(ws, rg, getTargetAddresses(ws, rg))

End Sub


Function getTargetAddresses(ws As Worksheet, rg As Range) As Object
    Dim dicTgtRng As Object
    
    Dim rg_src As Range

    Dim i As Long
    Dim j As Long
    
    Set dicTgtRng = CreateObject("Scripting.Dictionary")

    i = 1
    '列方向終了条件:
    '[セルの値が空白]かつ[セル結合していない]
    Do Until rg.Offset(0, i).Value = "" And Not rg.Offset(0, i).MergeCells
        DoEvents
        
        '表示列の場合のみ処理を実行する
        If Not rg.Offset(0, i).EntireColumn.Hidden Then
        
            Set rg_src = rg.Offset(0, i)
            
            Do
                DoEvents
                '次のセルが空白の時endで検索する
                If rg_src.Offset(1, 0).Value = "" Then
                    Set rg_src = rg_src.End(xlDown)
                '次のセルが空白でない時はひとつずつ確認する
                Else
                    j = 0
                    Do
                        DoEvents
                        j = j + 1
                    '数値のセルが見つかるか次のセルが空白の時はループを中断する
                    Loop Until IsNumeric(rg_src.Offset(j, 0)) Or rg_src.Offset(j + 1, 0).Value = ""
                    Set rg_src = rg_src.Offset(j, 0)
                End If
            
                'rg_src.Activate
            '行方向終了条件:
            '[セルの値が数値]または[最終行]
            Loop Until IsNumeric(rg_src.Value) Or rg_src.Row >= Rows.Count
            
            '[セルの値が数値]の時([最終行]でないとき)
            'セルアドレスを保存する
            If rg_src.Row <> Rows.Count Then
                If dicTgtRng.Exists(ws.Cells(rg_src.Row, rg.Column).Value) Then
                    dicTgtRng(ws.Cells(rg_src.Row, rg.Column).Value) = dicTgtRng(ws.Cells(rg_src.Row, rg.Column).Value) & "," & rg_src.Address
                Else
                    dicTgtRng.Add ws.Cells(rg_src.Row, rg.Column).Value, rg_src.Address
                End If
            End If
        End If
    
        i = i + 1
    Loop

    rg.Activate
    
    Set getTargetAddresses = dicTgtRng

End Function

Sub setColor(ws As Worksheet, rg As Range, dic As Object)
    Dim wsTmp As Worksheet
    
    Dim aryTgtRng() As String
    
    Dim clr() As Long

    Dim key As Variant

    Dim i As Long
    Dim j As Long
    
    
    '塗りつぶしの色を取得
    clr() = getColorAry


    '作業用ワークシートで日付を昇順に並べ替える(ディクショナリにキーを並べ替える機能なし)
    Set wsTmp = ThisWorkbook.Worksheets.Add
    wsTmp.Name = Format(Now(), "yymmdd-hhnnss")

    i = 0
    For Each key In dic.Keys
        DoEvents
        
        wsTmp.Range("A1").Offset(i, 0).Value = key
        wsTmp.Range("A1").Offset(i, 1).Value = dic(key)
        
        i = i + 1
    Next

    wsTmp.Range("A:B").Sort _
            Key1:=wsTmp.Columns(Range("A1").Column), _
            Order1:=xlAscending, Header:=xlNo


    i = 0
    Do Until wsTmp.Range("A1").Offset(i, 0).Value = ""
        DoEvents
        Debug.Print wsTmp.Range("A1").Offset(i, 1).Value
        aryTgtRng = Split(wsTmp.Range("A1").Offset(i, 1).Value, ",")
    
        '日付に色を塗る
        ws.Cells(ws.Range(aryTgtRng(0)).Row, rg.Column).Interior.Color = clr(i Mod (UBound(clr) + 1))
        For j = LBound(aryTgtRng) To UBound(aryTgtRng)
            ws.Range(aryTgtRng(j)).Interior.Color = clr(i Mod (UBound(clr) + 1))
            
            '見出しに色を塗る
            If ws.Cells(rg.Row, ws.Range(aryTgtRng(j)).Column).Interior.Color = 14277081 Then
                If ws.Cells(rg.Row, ws.Range(aryTgtRng(j)).Column).MergeCells Then
                    ws.Cells(rg.Row, ws.Range(aryTgtRng(j)).Column).MergeArea.Cells(1, 1).Interior.Color = clr(i Mod (UBound(clr) + 1))
                Else
                    ws.Cells(rg.Row, ws.Range(aryTgtRng(j)).Column).Interior.Color = clr(i Mod (UBound(clr) + 1))
                End If
            End If
        Next j
    
        i = i + 1
    Loop
    
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True

End Sub


Function getColorAry() As Variant
    Dim rg As Range
    Dim clr() As Long
    Dim clrIdx As Long
    
    Set rg = Worksheets("設定").Range("F4")
    
    
    ReDim clr(8)

    clrIdx = -1
    Do Until rg.Offset(clrIdx + 1, 0).Interior.Color = 16777215
        clrIdx = clrIdx + 1
        
        If clrIdx > UBound(clr) Then
            If UBound(clr) > 0 Then
                ReDim Preserve clr(UBound(clr) * 2)
            Else
                ReDim Preserve clr(8)
            End If
        End If
        
        clr(clrIdx) = rg.Offset(clrIdx, 0).Interior.Color
    Loop

    If clrIdx = -1 Then
        Err.Raise 1001, , "塗りつぶしの色が設定されていません"
    Else
        ReDim Preserve clr(clrIdx)
    End If

    getColorAry = clr

End Function
