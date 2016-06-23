Attribute VB_Name = "Module98"
Sub recurly_subs_period_pst()

'Changes timezone to PST
Dim lastRow As Long
lastRow = Range("A" & Rows.Count).End(xlUp).Row
 
Application.ScreenUpdating = False
Application.DisplayStatusBar = False

    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "period_ends_pst"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-14]-""09:00:00"""
    Range("AA2").Select
    Range("AA2").AutoFill Destination:=Range("AA2:AA" & lastRow)
    Columns("AA:AA").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("U:U").Select
    Selection.Replace What:="-0.375", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

