Attribute VB_Name = "Module23"
Sub dimp8_finreorg()
'
Dim msheet As Worksheet

Set msheet = ActiveSheet

    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("Q:Q").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:Q").Select
    Selection.ClearContents
End Sub

