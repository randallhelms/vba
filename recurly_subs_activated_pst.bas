Attribute VB_Name = "Module90"
Sub recurly_subs_activated_pst()

'Changes timezone to PST
Dim lastRow As Long
lastRow = Range("A" & Rows.Count).End(xlUp).Row
   
Application.ScreenUpdating = False
Application.DisplayStatusBar = False

    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "activated_at_pst"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]-""09:00:00"""
    Range("AA2").Select
    Range("AA2").AutoFill Destination:=Range("AA2:AA" & lastRow)
    Columns("AA:AA").Select
    Selection.Copy
    Columns("R:R").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub


