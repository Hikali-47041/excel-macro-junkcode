Sub sortContent()
    ' シート名は適切な名前に変更してください
    Const SHEETNAME As String = "Sheet1" 
    Const FLAGHEAD As String = "C1"
    Const DATEHEAD As String = "A1"
    Const HIDARIUE As String = "A1"
    Worksheets(SHEETNAME).sort.SortFields.Clear
    Worksheets(SHEETNAME).sort.SortFields.Add _
        Key:=Range(FLAGHEAD), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
    Worksheets(SHEETNAME).sort.SortFields.Add _
        Key:=Range(DATEHEAD), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
    With Worksheets(SHEETNAME).sort
        .SetRange Range(HIDARIUE, Range(HIDARIUE).SpecialCells(xlCellTypeLastCell))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
