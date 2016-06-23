Attribute VB_Name = "Module53"
Sub quick_fix_video_stuff()
Attribute quick_fix_video_stuff.VB_ProcData.VB_Invoke_Func = " \n14"
'
' quick_fix_video_stuff Macro
'
Dim msheet As Worksheet

Set msheet = ActiveSheet

    Columns("A:A").Select
    Selection.Replace What:="<n/a>", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "Watch Time"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]*RC[1])*1440"
    Range("D8").Select
    Selection.AutoFill Destination:=Range("D8:D61")
    Range("D8:D61").Select
    Range("D8").Select

End Sub
