Option Explicit

Public TargetCell As Range

'ボタンの値と色を変更する
Sub changeCaption(frm As Object, yearMonthStr As String)
    Dim Dt As Date
    Dim HolidaysDay As Object
    Dim targetYearMonth As String
    Dim wd As Long
    Dim ld As Long
    Dim i As Long
    
    frm.LblAlert.Caption = ""
    
    If Not isYearMonthStr(yearMonthStr) Then
        frm.LblAlert.Caption = "yyyymm形式で入力して下さい"
        frm.TxtYrMo.Text = frm.LblPrevYrMo.Caption
        Exit Sub
    End If
    
    Set HolidaysDay = getHolidaysDay(yearMonthStr)
    
    Dt = DateSerial(Left(yearMonthStr, 4), Right(yearMonthStr, 2), 1)
    wd = Weekday(Dt) - 1 ' 1=日曜, 2=月曜, ..., 7=土曜
    
    ld = Day(DateSerial(Year(Dt), Month(Dt) + 1, 0))
    
    For i = 1 To wd
        frm.Controls("btn" & i).Caption = ""
        frm.Controls("btn" & i).BackColor = vbButtonFace
        frm.Controls("btn" & i).Enabled = False
        Debug.Print "初期化->" & i
    Next
    
    For i = 1 To ld
        frm.Controls("btn" & i + wd).Caption = i
        frm.Controls("btn" & i + wd).BackColor = vbButtonFace
        
        If HolidaysDay Is Nothing Then
            If Weekday(DateSerial(Year(Dt), Month(Dt), i)) = 1 Then
               frm.Controls("btn" & i + wd).BackColor = rgbLightPink
            ElseIf Weekday(DateSerial(Year(Dt), Month(Dt), i)) = 7 Then
               frm.Controls("btn" & i + wd).BackColor = rgbLightSkyBlue
            End If
        Else
            If HolidaysDay.Exists(Format(i, "00")) Then
                    frm.Controls("btn" & i + wd).BackColor = rgbLightPink
            ElseIf Weekday(DateSerial(Year(Dt), Month(Dt), i)) = 1 Then
                frm.Controls("btn" & i + wd).BackColor = rgbLightPink
            ElseIf Weekday(DateSerial(Year(Dt), Month(Dt), i)) = 7 Then
                frm.Controls("btn" & i + wd).BackColor = rgbLightSkyBlue
            End If
        End If
        frm.Controls("btn" & i + wd).Enabled = True
        Debug.Print "設定中->" & (i + wd)
    Next

    For i = (ld + wd + 1) To 42
        frm.Controls("btn" & i).Caption = ""
        frm.Controls("btn" & i).BackColor = vbButtonFace
        frm.Controls("btn" & i).Enabled = False
        Debug.Print "初期化->" & i
    Next

    'yyyymm形式以外の文字列を入力されたときに直前の値に強制的に戻すために保持
    frm.LblPrevYrMo.Caption = Format(Dt, "yyyymm")

End Sub

Sub setDate(dateStr As String)
    TargetCell.Value = DateSerial(Left(dateStr, 4), Mid(dateStr, 5, 2), Right(dateStr, 2))
End Sub
