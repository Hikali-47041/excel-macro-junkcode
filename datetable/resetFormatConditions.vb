' XlRgbColor is BGR decimal
Sub resetFormatConditions()
    Application.StatusBar = "条件付き書式を再設定中"
    Cells.formatconditions.Delete
    Const SHEETNAME As String = "申し送り事項"
    Const TARGETCOLUMNS As String = "A:C"
    Dim fc As FormatCondition
    With Worksheets(SHEETNAME).Columns(TARGETCOLUMNS)
        ' important highlight #1
        Set fc = .formatconditions.Add(Type:=xlExpression, _
            Formula1:="=$C1=""重要""")
        With fc.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16048383 ' #ffe0f4
            .TintAndShade = 0
        End With
        fc.StopIfTrue = False

        ' odd line white
        Set fc = .formatconditions.Add(Type:=xlExpression, _
                Formula1:="=AND(ISODD(ROW($A1)), NOT($A1=""""))")
            With fc.Interior
                .PatternColorIndex = xlAutomatic
                .Color = 16579836 ' #fcfcfc
                .TintAndShade = 0
            End With
            fc.StopIfTrue = False

        ' even line light gray
        Set fc = .formatconditions.Add(Type:=xlExpression, _
                Formula1:="=AND(ISEVEN(ROW($A1)), NOT($A1=""""))")
            With fc.Interior
                .PatternColorIndex = xlAutomatic
                .Color = 15855855 ' #eff0f1 -> 15855855
                .TintAndShade = 0
            End With
            fc.StopIfTrue = False

        ' underline day diff
        Set fc = .formatconditions.Add(Type:=xlExpression, _
                Formula1:="=$A1<>$A2")
            With fc.Borders(xlBottom)
                .LineStyle = xlContinuous
                .TintAndShade = 0
                .Weight = xlThin
            End With
            fc.StopIfTrue = False
    End With
    Application.StatusBar = False
End Sub
