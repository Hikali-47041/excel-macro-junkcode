Sub button1_Click()
    ' URLがある行の値(数字)を入力(ここの値のみを変えてください)
    Const ROW As String = "2"
    ' URLが記載されているC3のセルの中身をクリップボードにコピーさせる(copyCellValueToClipboard)関数を呼び出し
    Call copyCellValueToClipboard("C" & ROW)
    ' ExcelのステータスバーにURLをコピーしたことを表示させる
    Application.StatusBar = "[" & Range("A" & ROW).Value & " " & Range("B" & ROW).Value & "] のURLをコピーしました!"
End Sub
