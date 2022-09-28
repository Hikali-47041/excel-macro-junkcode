' 引数 cellPlace で指定したセルの値をクリップボードにコピーする関数
Sub copyCellValueToClipboard(cellPlace As String)
    ' オブジェクト型の変数 ClipboardObject を定義
    Dim ClipboardObject As Object
    ' Clipboard: MSForms.DataObject のインスタンスを作成
    ' ref: https://qiita.com/cti1650/items/c0d2de73be45e4e4d10f
    Set ClipboardObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    With ClipboardObject
        ' cellPlace(対象)のセルをオブジェクトに格納
        .Settext Range(cellPlace).Value
        ' クリップボードに格納
        .PutInClipboard
    End With
End Sub
