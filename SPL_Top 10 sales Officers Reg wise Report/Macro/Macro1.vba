Sub Macro1()
'
' Macro1 Macro
'

'
    Range("B16").Select
    Sheets("Report").Select
    Columns("F:F").Select
    Range("F2").Activate
    ActiveWorkbook.Worksheets("Report").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Report").Sort.SortFields.Add Key:=Range("F2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Report").Sort
        .SetRange Range("A3:F445")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Range("A1:F1").Select
    ActiveWorkbook.Save
End Sub
