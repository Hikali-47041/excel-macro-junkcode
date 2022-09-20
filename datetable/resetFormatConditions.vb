' XlRgbColor is BGR decimal
Sub resetFormatConditions()
    Cells.FormatConditions.Delete
    Const SHEETNAME As String = "Sheet1"
    Const TARGETCOLUMNS As String = "A:C"
    Dim fc As FormatCondition
    With Worksheets(SHEETNAME).Columns(TARGETCOLUMNS)
        ' important highlight #1
        Set fc = .FormatConditions.Add (Type:=xlExpression, _
            Formula1:= "=$C1=1") 
        With fc.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16048383 ' #ffe0f4
            .TintAndShade = 0
        End With
        fc.StopIfTrue = False
        
        ' important highlight #2
        Set fc = .FormatConditions.Add (Type:=xlExpression, _
            Formula1:= "=$C1=2") 
        With fc.Interior
            .PatternColorIndex = xlAutomatic
            .Color = 15385599 ' #ffc3ea
            .TintAndShade = 0
        End With
        fc.StopIfTrue = False

        ' odd line white
        Set fc = .FormatConditions.Add (Type:=xlExpression, _
                Formula1:= "=AND(ISODD(ROW($A1)), NOT($A1=""""))")
            With fc.Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            fc.StopIfTrue = False

        ' even line light gray
        Set fc = .FormatConditions.Add (Type:=xlExpression, _
                Formula1:= "=AND(ISEVEN(ROW($A1)), NOT($A1=""""))") 
            With fc.Interior
                .PatternColorIndex = xlAutomatic
                .Color = 15855855 ' #eff0f1 -> 15855855
                .TintAndShade = 0
            End With
            fc.StopIfTrue = False

        ' underline day diff
        Set fc = .FormatConditions.Add (Type:=xlExpression, _
                Formula1:="=$A1<>$A2")
            With fc.Borders(xlBottom)
                .LineStyle = xlContinuous
                .TintAndShade = 0
                .Weight = xlThin
            End With
            fc.StopIfTrue = False
    End With
End Sub
