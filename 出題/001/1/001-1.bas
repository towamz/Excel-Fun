Option Explicit

Sub task001_1()
    Dim ws As Worksheet
    Dim rg As Range
    Dim rg_src As Range
    
    Dim clr As Long
    Dim i As Long
    Dim j As Long

    clr = Worksheets("設定").Range("D4").Interior.Color

    Set ws = Worksheets("スケジュール")
    Set rg = ws.Range("B3")


    i = 1
    '列方向終了条件:
    '[セルの値が空白]かつ[セル結合していない]
    Do Until rg.Offset(0, i).Value = "" And Not rg.Offset(0, i).MergeCells
        DoEvents
        
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
        
        '[セルの値が数値]の時
        'セルの色の塗りつぶしを実施する
        If IsNumeric(rg_src.Value) Then
            If rg.Offset(0, i).MergeCells Then
                Debug.Print rg.Offset(0, i).MergeArea.Cells(1, 1).Value
                rg.Offset(0, i).MergeArea.Cells(1, 1).Interior.Color = clr
            
            Else
                Debug.Print rg.Offset(0, i).Value
                rg.Offset(0, i).Interior.Color = clr
            End If
        
            Debug.Print rg_src.Address
            Debug.Print Cells(rg_src.Row, rg.Column).Address
            rg_src.Interior.Color = clr
            Cells(rg_src.Row, rg.Column).Interior.Color = clr
        End If
    
        i = i + 1
    Loop

    rg.Activate

End Sub
