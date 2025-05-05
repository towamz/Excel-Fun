Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Application.Intersect(Target, Range("A2:O2")) Is Nothing Then
        Application.EnableEvents = False
        Call getFilteredData
        Application.EnableEvents = True
    End If
End Sub
