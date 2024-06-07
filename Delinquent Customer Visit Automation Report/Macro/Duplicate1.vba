Sub Macro1()
'
' Macro1 Macro
'

'
    Sheets("Today Report").Select
    Columns("K:K").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Columns( _
        "Y:Y"), Unique:=True
    Columns("Y:Y").Select
    Selection.Cut
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Application.CutCopyMode = False
    Sheets("Sheet1").Move Before:=Sheets(1)
    Range("A1").Select
    Sheets("Today Report").Select
    Columns("Y:Y").Select
    Selection.Cut
    Sheets("Sheet1").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Summary"
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
