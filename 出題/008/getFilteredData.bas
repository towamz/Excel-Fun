Option Explicit

Enum searchFields
    Name = 1
    Furigana
    MailAddress
    Department
    Title
    Sex
    Age
    Birthday
    MaritalStatus
    BloodType
    Prefecture
    LandlinePhoneNumber
    MobilePhoneNumber
    MobileCarrier
    CurryEatingStyles
    MaxNumber_NotForUse
End Enum

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
    
    For i = 1 To searchFields.MaxNumber_NotForUse - 1
        Select Case i
            Case searchFields.Name
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Furigana
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.MailAddress
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Department
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Title
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Sex       '検索キー１つだけの関数を作るか?
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Age
                Call autoFilterNumber(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Birthday
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.MaritalStatus     '検索キー１つだけの関数を作るか?
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.BloodType
                Call autoFilterBloodType(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.Prefecture
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.LandlinePhoneNumber
                Call autoFilterPhoneNumber(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.MobilePhoneNumber
                Call autoFilterPhoneNumber(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.MobileCarrier     '検索キー１つだけの関数を作るか?
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
            Case searchFields.CurryEatingStyles
                Call autoFilterString(rgOrig, i, wsDest.Cells(2, i).Value)
        End Select
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


Sub autoFilterString(rg As Range, colIndex As Long, val As String)
    Dim REX As Object
    Dim aryVal() As String
    
    '正規表現オブジェクト
    Set REX = CreateObject("VBScript.RegExp")
    REX.IgnoreCase = True
    REX.Global = True
    REX.Pattern = "(\s|　)+"
    
    aryVal = Split(REX.Replace(Trim(val), Chr(32)), Chr(32))
    Select Case UBound(aryVal)
        Case -1
            '空白なので処理しない
        Case 0
            rg.AutoFilter Field:=colIndex, Criteria1:="*" & aryVal(0) & "*"
        Case Else
            rg.AutoFilter Field:=colIndex, Criteria1:="*" & aryVal(0) & "*", Operator:=xlAnd, Criteria2:="*" & aryVal(1) & "*"
    End Select
    
    If UBound(aryVal) > 1 Then
        Debug.Print "検索キーは2つまでです:" & colIndex
    End If
End Sub

Sub autoFilterNumber(rg As Range, colIndex As Long, val As String)
    Dim REX As Object
    Dim aryVal() As String
    
    '正規表現オブジェクト
    Set REX = CreateObject("VBScript.RegExp")
    REX.IgnoreCase = True
    REX.Global = True
    REX.Pattern = "(\s|　)+"

    aryVal = Split(REX.Replace(Trim(val), Chr(32)), Chr(32))
    Select Case UBound(aryVal)
        Case -1
            '空白なので処理しない
        Case 0
            rg.AutoFilter Field:=colIndex, Criteria1:=aryVal(0)
        Case 1
            '数値以外指定の動作
            'if IsNumericを有効化:全件表示
            '              無効化:0件
'            If IsNumeric(aryVal(0)) And IsNumeric(aryVal(1)) Then
                If aryVal(0) < aryVal(1) Then
                    rg.AutoFilter Field:=colIndex, Criteria1:=">=" & aryVal(0), Operator:=xlAnd, Criteria2:="<=" & aryVal(1)
                Else
                    rg.AutoFilter Field:=colIndex, Criteria1:=">=" & aryVal(1), Operator:=xlAnd, Criteria2:="<=" & aryVal(0)
                End If
'            End If
        Case Else
            rg.AutoFilter Field:=colIndex, Criteria1:=aryVal, Operator:=xlFilterValues
    
    End Select

    If UBound(aryVal) > 1 Then
        Debug.Print "検索キーは2つまでです:" & colIndex
    End If
End Sub

Sub autoFilterBloodType(rg As Range, colIndex As Long, val As String)
    Dim REX As Object
    Dim aryVal() As String
    
    '正規表現オブジェクト
    Set REX = CreateObject("VBScript.RegExp")
    REX.IgnoreCase = True
    REX.Global = True
    REX.Pattern = "(\s|　)+"
    
    aryVal = Split(REX.Replace(Trim(val), Chr(32)), Chr(32))
    Select Case UBound(aryVal)
        Case -1
            '空白なので処理しない
        Case Else
            'A型とAB型がB型が同時検索されないようにワイルドカードを使わない
            If Right(aryVal(0), 1) = "型" Then
                rg.AutoFilter Field:=colIndex, Criteria1:=aryVal(0)
            
            Else
                rg.AutoFilter Field:=colIndex, Criteria1:=aryVal(0) & "型"
            
            End If
    End Select
    
    If UBound(aryVal) > 0 Then
        Debug.Print "検索キーは1つまでです:" & colIndex
    End If
End Sub


Sub autoFilterPhoneNumber(rg As Range, colIndex As Long, val As String)
    Dim REX As Object
    Dim aryVal() As String
    
    '正規表現オブジェクト
    Set REX = CreateObject("VBScript.RegExp")
    REX.IgnoreCase = True
    REX.Global = True
    REX.Pattern = "(\s|　)+"
    
    aryVal = Split(REX.Replace(Trim(val), ""), "-")
    Select Case UBound(aryVal)
        Case -1
            '空白なので処理しない
        Case 0
            rg.AutoFilter Field:=colIndex, Criteria1:="*" & aryVal(0) & "*"
        Case 1
            rg.AutoFilter Field:=colIndex, Criteria1:="*" & aryVal(0) & "*-*" & aryVal(1) & "*-*"
            Debug.Print "*" & aryVal(0) & "*-*" & aryVal(1) & "*-"
        Case Else
            rg.AutoFilter Field:=colIndex, Criteria1:="*" & aryVal(0) & "*-*" & aryVal(1) & "*-*" & aryVal(2) & "*"
            Debug.Print "*" & aryVal(0) & "*-*" & aryVal(1) & "*-*" & aryVal(2) & "*"

    End Select
    
    If UBound(aryVal) > 2 Then
        Debug.Print "検索キーは2つまでです:" & colIndex
    End If
End Sub