Attribute VB_Name = "Module39"
Sub dimp_newsheet()
'Creates a new sheet and then runs a match formula

Dim msheet As Worksheet
Dim lastRow As Long
Dim wb1 As Window
lastRow = Range("A" & Rows.Count).End(xlUp).Row

Set msheet = ActiveSheet
Set wb1 = ActiveWindow

' Paste from Previous Section

    Columns("A:A").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste

' Copy to Column B

    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("C1").Select
    ActiveCell = "Match"
        
' Copy from existing Title Info

    Application.WindowState = xlNormal
    Windows("Title Info (Current).xlsx").Activate
    Columns("A:A").Select
    Selection.Copy
    wb1.Activate
    Range("E1").Select
    ActiveSheet.Paste
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
' remove spaces

    Range("E:E,A:B").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
' Add formulas

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MATCH(RC[-1],R2C5:R26000C5,0)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("c2:c" & lastRow)
    
' Copy and paste

    Range("C1").Select
    Range("C1:C" & lastRow).Copy
    
    msheet.Activate
    Columns("S").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
End Sub
