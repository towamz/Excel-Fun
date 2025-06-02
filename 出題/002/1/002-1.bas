Sub task002_1()
    Dim rng As Range
    Dim rngDiff As Range
    Dim rgbOrig As Long
    Dim rgb As Variant
    Dim rOrig As Long, gOrig As Long, bOrig As Long
    Dim min As Long
    Dim r As Long, g As Long, b As Long
    Dim i As Long
    
    Set rng = Worksheets("VBAで使えるカラー定数一覧").Range("B2")
    Set rngDiff = Worksheets("実行").Range("E2")

    '指定色のRGBを取得
    rgbOrig = Worksheets("実行").Range("B1").Interior.Color
    rOrig = rgbOrig Mod 256
    gOrig = (rgbOrig / 256) Mod 256
    bOrig = (rgbOrig / 256 / 256) Mod 256
    
    Worksheets("実行").Range("B2").Value = rOrig & "," & gOrig & "," & bOrig

    
    '指定色と定数のRGBの差分の絶対値を取得
    i = 0
    Do Until rng.Offset(i, 0).Value = ""
        DoEvents

        rgb = Split(rng.Offset(i, 0).Value, ",")
        r = CLng(rgb(0))
        g = CLng(rgb(1))
        b = CLng(rgb(2))
        
        rngDiff.Offset(i, 0) = Abs(rOrig - r)
        rngDiff.Offset(i, 1) = Abs(gOrig - g)
        rngDiff.Offset(i, 2) = Abs(bOrig - b)
        
        i = i + 1
    Loop

    '差分の合計を取得
    Worksheets("実行").Range(rngDiff.Offset(0, 3), rngDiff.Offset(i - 1, 3)).Formula = "=SUM(E2:G2)"
    
    '差分の一番小さい値を取得
    min = WorksheetFunction.min(Worksheets("実行").Range(rngDiff.Offset(0, 3), rngDiff.Offset(i - 1, 3)))
    
    '差分が最小の色の行を検索し見つかったら結果に書き出してexitする
    i = 2
    Do Until Worksheets("実行").Cells(i, 8).Value = ""
        DoEvents

        If Worksheets("実行").Cells(i, 8).Value = min Then
            Worksheets("実行").Range("B5").Interior.Color = Worksheets("VBAで使えるカラー定数一覧").Range("A" & i).Interior.Color
            Worksheets("実行").Range("B6").Value = Worksheets("VBAで使えるカラー定数一覧").Range("B" & i).Value
            Worksheets("実行").Range("B7").Value = Worksheets("VBAで使えるカラー定数一覧").Range("C" & i).Value
            Worksheets("実行").Range("B8").Value = Worksheets("VBAで使えるカラー定数一覧").Range("D" & i).Value
            Exit Do
        End If
        i = i + 1
    Loop

End Sub
