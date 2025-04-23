Sub makeBills2()
    Dim MKB As New ClsMakeBills2

    Application.ScreenUpdating = False
    Call MKB.makeBills2
    Application.ScreenUpdating = True

    MsgBox "終了しました"

End Sub
