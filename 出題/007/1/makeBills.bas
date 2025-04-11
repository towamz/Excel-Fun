Sub makeBills()
    Dim GFD As ClsGetFilteredData
    Dim GDR As ClsGetDataRows
    Dim PDF As ClsGetPdf
    Dim testRange As Range
    Dim ShData As Worksheet
    Dim ShDetailBills() As Worksheet
    Dim ShTotalBill As Worksheet
    Dim detailBillsName As String
    Dim totalBillName As String
    Dim detailBillsCount As Long

    Dim workingColumnStartEntire As Range
    Dim workingColumnStartHeader As Range
    Dim detailBillsColumnStartEntire As Range
    Dim detailBillsColumnStartHeader As Range
    
    Dim companyName As String
    Dim targetYear As Long
    
    Dim i As Long
    
    Const productCodesColumnStr As String = "F"
    Const salesColumnStr As String = "J"
    Const detailBillsColumnStartStr As String = "K"
    Const workingColumnStartStr As String = "Q"

    totalBillName = "請求書"
    detailBillsName = "請求明細書"

    Set PDF = New ClsGetPdf
    Set GFD = New ClsGetFilteredData

    Set GFD.setBook = ThisWorkbook
    GFD.SheetName = "販売データ"
    GFD.HeaderRangeName = "B4:K4"
    GFD.TargetColumnLetter = "B"
    GFD.CompanyCode = Worksheets("実行シート").Range("C2").Value
    GFD.TargetMonth = Worksheets("実行シート").Range("C5").Value
    companyName = Worksheets("実行シート").Range("C3").Value
    targetYear = Worksheets("実行シート").Range("C4").Value

    Set workingColumnStartEntire = ShData.Range(workingColumnStartStr & ":" & workingColumnStartStr)
    Set workingColumnStartHeader = ShData.Range(workingColumnStartStr & "1")
    Set detailBillsColumnStartEntire = ShData.Range(detailBillsColumnStartStr & ":" & detailBillsColumnStartStr)
    Set detailBillsColumnStartHeader = ShData.Range(detailBillsColumnStartStr & "1")
    
    Set ShData = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ShData.Name = "data" & Format(Now(), "yymmdd-hhnnss")
    
    
    Set testRange = GFD.getFilteredData
    testRange.Copy
    ShData.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    If ShData.Range("A2") = "" Then
        Err.Raise 1090, , "指定の取引先コード・年月に販売データがありません"
    End If
    
    '請求書明細コピー用列を作成
    'ご購入日(D列=月, E列=日), 商品(G列),単価(H列),数量(I列),金額(J列)
    ShData.Range(Split(detailBillsColumnStartHeader.Address, "$")(1) & 1).Value = "No"
    
    
    Set GDR = Nothing
    Set GDR = New ClsGetDataRows
    Set GDR.ws = ShData
'    GDR.TargetColumnLetter = "O"
'    shData.Range("P1").Value = "種別レコード数"
'    shData.Range("P2:P" & GDR.getLastRow).Formula = "=SUMIF(N:N,O2,L:L)"

    GDR.TargetColumnLetter = "A"
    
    detailBillsColumnStartHeader.Offset(1, 0).Value = 1
    detailBillsColumnStartHeader.Offset(2, 0).Value = 2
    ShData.Range(detailBillsColumnStartHeader.Offset(1, 0), detailBillsColumnStartHeader.Offset(2, 0)).AutoFill _
    Destination:=ShData.Range(detailBillsColumnStartHeader.Offset(1, 0), detailBillsColumnStartHeader.Offset(GDR.getLastRow - 1, 0))
    
    
    ShData.Range(Split(detailBillsColumnStartHeader.Offset(0, 1).Address, "$")(1) & 1).Value = "ご購入日"
    i = 2
    Do Until ShData.Range("D" & i).Value = ""
        ShData.Range(Split(detailBillsColumnStartHeader.Offset(0, 1).Address, "$")(1) & i).Value = ShData.Range("D" & i).Value & "/" & ShData.Range("E" & i).Value
        ShData.Range(Split(detailBillsColumnStartHeader.Offset(0, 1).Address, "$")(1) & i).NumberFormat = "m/d"
        i = i + 1
    Loop
    
    
    detailBillsColumnStartEntire.Offset(0, 2).Value = ShData.Range("G:G").Value
    detailBillsColumnStartHeader.Offset(0, 2).Value = "商品"
    
    detailBillsColumnStartEntire.Offset(0, 3).Value = ShData.Range("H:H").Value
    detailBillsColumnStartHeader.Offset(0, 3).Value = "単価"
    
    detailBillsColumnStartEntire.Offset(0, 4).Value = ShData.Range("I:I").Value
    detailBillsColumnStartHeader.Offset(0, 4).Value = "数量"
    
    detailBillsColumnStartEntire.Offset(0, 5).Value = ShData.Range("J:J").Value
    detailBillsColumnStartHeader.Offset(0, 5).Value = "金額"
    
    
    
    '重複データ削除取得(F列=商品コード->K列)
'    shData.Range("F:F").Copy Destination:=shData.Range("K:K")
'    shData.Range("K:K").RemoveDuplicates Columns:=1, Header:=xlYes
'    shData.Range("K:K").Sort Key1:=shData.Range("K:K"), Order1:=xlAscending, Header:=xlYes
'    shData.Range("K1").Value = "商品コード(重複なし)"


        ShData.Range(productCodesColumnStr & ":" & productCodesColumnStr).Copy Destination:=workingColumnStartEntire
    workingColumnStartEntire.RemoveDuplicates Columns:=1, Header:=xlYes
    workingColumnStartEntire.Sort Key1:=ShData.Range(workingColumnStartEntire.Address(False, False)), Order1:=xlAscending, Header:=xlYes
    workingColumnStartHeader.Value = "商品コード(重複なし)"




    '商品コード別件数
    Set GDR = Nothing
    Set GDR = New ClsGetDataRows
    Set GDR.ws = ShData
'    GDR.TargetColumnLetter = "K"
'    shData.Range("L1").Value = "商品コード別レコード数"
'    shData.Range("L2:L" & GDR.getLastRow).Formula = "=COUNTIF(F:F,K2)"
    
    GDR.TargetColumnLetter = workingColumnStartEntire.Address
    workingColumnStartHeader.Offset(0, 1).Value = "商品コード別レコード数"
'    shData.Range(workingColumnStartHeader.Offset(1, 1), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 1)).Formula = "=COUNTIF(F:F,K2)"
    ShData.Range(workingColumnStartHeader.Offset(1, 1), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 1)).Formula = "=COUNTIF(" & productCodesColumnStr & ":" & productCodesColumnStr & "," & workingColumnStartHeader.Offset(1, 0).Address(False, False) & ")"
    

    
    
    '商品コード別売上
'    shData.Range("M1").Value = "商品コード別売上"
'    shData.Range("M2:M" & GDR.getLastRow).Formula = "=SUMIF(F:F,K2,J:J)"
'    shData.Range(workingColumnStartHeader.Offset(1, 2), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 2)).Formula = "=SUMIF(F:F,K2,J:J)"
    workingColumnStartHeader.Offset(0, 2).Value = "商品コード別売上"
    ShData.Range(workingColumnStartHeader.Offset(1, 2), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 2)).Formula = "=SUMIF(" & productCodesColumnStr & ":" & productCodesColumnStr & "," & workingColumnStartHeader.Offset(1, 0).Address(False, False) & "," & salesColumnStr & ":" & salesColumnStr & ")"
    


    '種別
'    shData.Range("N1").Value = "種別"
'    shData.Range("N2:N" & GDR.getLastRow).Formula = "=XLOOKUP(K2,商品マスタ!A:A,商品マスタ!C:C)"
    
    workingColumnStartHeader.Offset(0, 3).Value = "種別(重複あり)"
'    shData.Range(workingColumnStartHeader.Offset(1, 3), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 3)).Formula = "=XLOOKUP(K2,商品マスタ!A:A,商品マスタ!C:C)"
    ShData.Range(workingColumnStartHeader.Offset(1, 3), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 3)).Formula = "=XLOOKUP(" & workingColumnStartHeader.Offset(1, 0).Address(False, False) & ",商品マスタ!A:A,商品マスタ!C:C)"
    
    
    '重複データ削除取得(M列=種別->N列)
'    shData.Range("O:O").Value = shData.Range("N:N").Value
'    shData.Range("O:O").RemoveDuplicates Columns:=1, Header:=xlYes
'    shData.Range("O:O").Sort Key1:=shData.Range("O:O"), Order1:=xlAscending, Header:=xlYes
'    shData.Range("O1").Value = "種別(重複なし)"
    
    workingColumnStartEntire.Offset(0, 5).Value = workingColumnStartEntire.Offset(0, 3).Value
    workingColumnStartEntire.Offset(0, 5).RemoveDuplicates Columns:=1, Header:=xlYes
    workingColumnStartEntire.Offset(0, 5).Sort Key1:=workingColumnStartEntire.Offset(0, 5), Order1:=xlAscending, Header:=xlYes
    workingColumnStartHeader.Offset(0, 5).Value = "種別"
    
    
'   'No
    workingColumnStartHeader.Offset(0, 4).Value = "No"
    
    i = 1
    Do Until workingColumnStartHeader.Offset(i, 5).Value = ""
        workingColumnStartHeader.Offset(i, 4).Value = i
        i = i + 1
    Loop
    
    '種別レコード数
    Set GDR = Nothing
    Set GDR = New ClsGetDataRows
    Set GDR.ws = ShData
'    GDR.TargetColumnLetter = "O"
'    shData.Range("P1").Value = "種別レコード数"
'    shData.Range("P2:P" & GDR.getLastRow).Formula = "=SUMIF(N:N,O2,L:L)"

    GDR.TargetColumnLetter = workingColumnStartEntire.Offset(0, 5).Address
    workingColumnStartHeader.Offset(0, 6).Value = "件数"
'    shData.Range(workingColumnStartHeader.Offset(1, 5), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 5)).Formula = "=SUMIF(N:N,O2,L:L)"
    ShData.Range(workingColumnStartHeader.Offset(1, 6), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 6)).Formula = _
    "=SUMIF(" & workingColumnStartEntire.Offset(0, 3).Address(False, False) & "," & workingColumnStartHeader.Offset(1, 5).Address(False, False) & "," & workingColumnStartEntire.Offset(0, 1).Address(False, False) & ")"



    '種別レコード数
'    shData.Range("Q1").Value = "種別売り上げ"
'    shData.Range("Q2:Q" & GDR.getLastRow).Formula = "=SUMIF(N:N,O2,M:M)"

    workingColumnStartHeader.Offset(0, 7).Value = "金額"
'    shData.Range(workingColumnStartHeader.Offset(1, 6), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 6)).Formula = "=SUMIF(N:N,O2,M:M)"
    ShData.Range(workingColumnStartHeader.Offset(1, 7), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 7)).Formula = "=SUMIF(" & workingColumnStartEntire.Offset(0, 3).Address(False, False) & "," & workingColumnStartHeader.Offset(1, 5).Address(False, False) & "," & workingColumnStartEntire.Offset(0, 2).Address(False, False) & ")"

    workingColumnStartHeader.Offset(0, 8).Value = "合計"
'    shData.Range(workingColumnStartHeader.Offset(1, 6), workingColumnStartHeader.Offset(GDR.getLastRow - 1, 6)).Formula = "=SUM(V2:V3)"
    workingColumnStartHeader.Offset(1, 8).Formula = _
    "=SUM(" & workingColumnStartHeader.Offset(1, 7).Address(False, False) & "," & workingColumnStartHeader.Offset(GDR.getLastRow - 1, 7).Address(False, False) & ")"

    workingColumnStartHeader.Offset(2, 8).Formula = _
    "=Int(" & workingColumnStartHeader.Offset(1, 8).Address(False, False) & "* 0.1)"

    workingColumnStartHeader.Offset(3, 8).Formula = _
    "=" & workingColumnStartHeader.Offset(1, 8).Address(False, False) & "+" & workingColumnStartHeader.Offset(2, 8).Address(False, False)




    '請求書作成
    Sheets(totalBillName).Copy After:=Sheets(Sheets.Count)
    Set ShTotalBill = ActiveSheet
    ShTotalBill.Name = totalBillName & "(1)"
    PDF.WsName = ShTotalBill.Name

    ShTotalBill.Range("H2").Value = Now()
    ShTotalBill.Range("H2").NumberFormat = "ggge年mm月dd日"
    
    ShTotalBill.Range("B6").Value = ShData.Range("C2").Value

    ShTotalBill.Range("B8").Value = "件名：" & Format(DateSerial(targetYear, GFD.TargetMonth, 1), "yyyy年m月分") & "について"

    ShTotalBill.Range("D11").Value = workingColumnStartHeader.Offset(3, 8).Value

    '■■■■■■仕様確認■■■■■
    'shTotalBill.Range("D13").Value = DateSerial(targetYear, GFD.TargetMonth + 2, 0)
    ShTotalBill.Range("D13").Value = Format(DateSerial(Year(Now()), month(Now()) + 2, 0))



    Set GDR = Nothing
    Set GDR = New ClsGetDataRows
    Set GDR.ws = ShData
    GDR.TargetColumnLetter = workingColumnStartEntire.Offset(0, 4).Address
'    shTotalBill.Range("B16:").Value = a
'    shTotalBill.Range(shTotalBill.Range("B16"), shTotalBill.Range("B16").Offset(GDR.getLastRow - 15, 0)).Formula = _
'    "=SUMIF(" & workingColumnStartEntire.Offset(0, 3).Address(False, False) & "," & workingColumnStartHeader.Offset(1, 5).Address(False, False) & "," & workingColumnStartEntire.Offset(0, 1).Address(False, False) & ")"
'
    ShTotalBill.Range("B16:B20").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 4), workingColumnStartHeader.Offset(5, 4)).Value

    ShTotalBill.Range("C16:C20").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 5), workingColumnStartHeader.Offset(5, 5)).Value

    ShTotalBill.Range("G16:G20").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 6), workingColumnStartHeader.Offset(5, 6)).Value

    ShTotalBill.Range("I16:I20").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 7), workingColumnStartHeader.Offset(5, 7)).Value

    ShTotalBill.Range("I21:I23").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 8), workingColumnStartHeader.Offset(5, 8)).Value

    ShTotalBill.Range("B26").Value = _
    "対象取引期間：" & Format(DateSerial(targetYear, GFD.TargetMonth, 1), "yyyy/m/d") & "～" & Format(DateSerial(targetYear, GFD.TargetMonth + 1, 0), "yyyy/m/d")






    '請求明細書作成

    Set GDR = Nothing
    Set GDR = New ClsGetDataRows
    Set GDR.ws = ShData
    GDR.TargetColumnLetter = "A"
    
    '請求書の枚数を取得(1シート20件)
    detailBillsCount = Int((GDR.getLastRow - 2) / 20)

    ReDim ShDetailBills(detailBillsCount)
    
    
    For i = 0 To UBound(ShDetailBills)
    
        Sheets(detailBillsName).Copy After:=Sheets(Sheets.Count)
        Set ShDetailBills(i) = ActiveSheet
        ShDetailBills(i).Name = detailBillsName & "(" & (i + 1) & ")"
        PDF.WsName = ShDetailBills(i).Name
        
        ShDetailBills(i).Range("B5").Value = ShData.Range("C2").Value

        ShDetailBills(i).Range(ShDetailBills(i).Range("B8").Offset(0, 0), ShDetailBills(i).Range("B8").Offset(19, 0)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 0), detailBillsColumnStartHeader.Offset(20 + i * 20, 0)).Value


        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 0), ShDetailBills(i).Range("C8").Offset(19, 0)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 1), detailBillsColumnStartHeader.Offset(20 + i * 20, 1)).Value
        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 0), ShDetailBills(i).Range("C8").Offset(19, 0)).NumberFormat = _
        "m/d"
        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 1), ShDetailBills(i).Range("C8").Offset(19, 1)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 2), detailBillsColumnStartHeader.Offset(20 + i * 20, 2)).Value

        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 4), ShDetailBills(i).Range("C8").Offset(19, 4)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 3), detailBillsColumnStartHeader.Offset(20 + i * 20, 3)).Value
        
        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 6), ShDetailBills(i).Range("C8").Offset(19, 6)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 4), detailBillsColumnStartHeader.Offset(20 + i * 20, 4)).Value
    
        ShDetailBills(i).Range(ShDetailBills(i).Range("C8").Offset(0, 7), ShDetailBills(i).Range("C8").Offset(19, 7)).Value = _
        ShData.Range(detailBillsColumnStartHeader.Offset(1 + i * 20, 5), detailBillsColumnStartHeader.Offset(20 + i * 20, 5)).Value
    
    Next

    ShDetailBills(UBound(ShDetailBills)).Range("K28:K30").Value = _
    ShData.Range(workingColumnStartHeader.Offset(1, 8), workingColumnStartHeader.Offset(3, 8)).Value

    'PDF作成
    PDF.TargetDirectory = ThisWorkbook.Path & "\請求書"
    PDF.PdfName = "請求書" & Format(DateSerial(targetYear, GFD.TargetMonth, 1), "yyyy年m月") & "(" & companyName & "様).pdf"

    PDF.savePDF

    PDF.IsAlertForDeleteSheet = False
    PDF.deleteSheets
    
    Application.DisplayAlerts = False
    ShData.Delete
    Application.DisplayAlerts = True

End Sub


