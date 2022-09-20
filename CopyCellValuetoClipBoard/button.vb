Sub Button1_Click()
    Const ROW As Integer = 2
    Call CopyCellPlaceValueToClipboard("C" & ROW)
    Application.StatusBar = "Copied: " & Range("A" & ROW) & " " & Range("B" & ROW)
End Sub
