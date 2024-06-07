Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Region1").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Region1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Region1").Sort.SortFields.Add Key:=Range("F1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Region1").Sort
        .SetRange Range("A2:F300")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Region2").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Region2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Region2").Sort.SortFields.Add Key:=Range("F1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Region2").Sort
        .SetRange Range("A2:F300")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Region3").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Region3").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Region3").Sort.SortFields.Add Key:=Range("F1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Region3").Sort
        .SetRange Range("A2:F300")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Sheets("Region4").Select
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Region4").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Region4").Sort.SortFields.Add Key:=Range("F1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Region4").Sort
        .SetRange Range("A2:F300")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G2").Select
    Sheets("Region2").Select
    Range("G2").Select
    Sheets("Region1").Select
    Range("G2").Select
    ActiveWorkbook.Save
End Sub
