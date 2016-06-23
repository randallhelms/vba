Attribute VB_Name = "Module21"
Sub dimp6_matching()
'
' dimp6_matching Macro
'
Dim msheet As Worksheet

Set msheet = ActiveSheet

 Columns("A:A").Select
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Match"
    Range("C2").Select
    Windows("Title Info 2015.06.30.xlsx").Activate
    Selection.Copy
    Windows("Book1").Activate
    Range("E1").Select
    ActiveSheet.Paste
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MATCH(RC[-1],R2C5:R20697C5,0)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C21086")
    Range("C2:C21086").Select
    Range("E1").Select
End Sub

