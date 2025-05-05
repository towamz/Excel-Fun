Option Explicit
Public Sub getFilteredData()
    Dim wsOrig As Worksheet
    Dim wsDest As Worksheet
    Dim rgOrig As Range
    Dim rgDest As Range
    Dim i As Long
    
'    '他でも流用できるようにローカルオブジェクト変数に再代入する
    Set wsOrig = Worksheets("社員名簿")
    Set wsDest = Worksheets("検索")
    
    '現在表示されている値をクリアする
    wsDest.Rows("3:" & Rows.Count).Clear
    
    
    'オートフィルタを解除する
    On Error Resume Next
    wsOrig.ShowAllData 'フィルタが設定されていないとエラー発生
    On Error GoTo 0
    wsOrig.AutoFilterMode = False
    
    
'    'データのあるセル
    Set rgOrig = wsOrig.Range("A1").CurrentRegion
    
    
    '■■■■■抽出■■■■■
    rgOrig.AutoFilter
    
    For i = 1 To 15
        If wsDest.Cells(2, i).Value <> "" Then
            rgOrig.AutoFilter Field:=i, Criteria1:="*" & wsDest.Cells(2, i).Value & "*"
        End If
    Next i
    
    On Error Resume Next
    'タイトルを含みコピー
    Set rgDest = rgOrig.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    'データがない場合でもタイトルのみコピー
    rgDest.Copy
    wsDest.Range("A3").PasteSpecial Paste:=xlPasteValues


    'オートフィルタを解除する
    On Error Resume Next
    wsOrig.ShowAllData 'フィルタが設定されていないとエラー発生
    On Error GoTo 0
    wsOrig.AutoFilterMode = False

End Sub

