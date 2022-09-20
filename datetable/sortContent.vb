Sub sortContent()
    Const SHEETNAME As String = "Sheet1"
    Const FLAGHEAD As String = "C1"
    Const DATEHEAD As String = "A1"
    Const HIDARIUE As String = "A1"
    Worksheets(SHEETNAME).Sort.SortFields.Clear
    Worksheets(SHEETNAME).Sort.SortFields.Add2 _
        Key:=Range(FLAGHEAD), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
    Worksheets(SHEETNAME).Sort.SortFields.Add2 _
        Key:=Range(DATEHEAD), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
    With Worksheets(SHEETNAME).Sort
        .SetRange Range(HIDARIUE, Range(HIDARIUE).SpecialCells(xlCellTypeLastCell))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
