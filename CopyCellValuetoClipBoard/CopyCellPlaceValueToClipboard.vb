Sub CopyCellPlaceValueToClipboard(ByVal cellPlace As String)
    Dim clipboardObject As Object
    Set clipboardObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    With clipboardObject
        .SetText Range(cellPlace).Value
        .PutInClipboard
    End With
End Sub
