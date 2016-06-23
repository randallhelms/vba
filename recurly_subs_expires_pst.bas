Attribute VB_Name = "Module91"
Sub recurly_subs_expires_pst()

'Changes timezone to PST
Dim lastRow As Long
lastRow = Range("A" & Rows.Count).End(xlUp).Row
   
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "expires_at_pst"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]-""09:00:00"""
    Range("AA2").Select
    Range("AA2").AutoFill Destination:=Range("AA2:AA" & lastRow)
    Columns("AA:AA").Select
    Selection.Copy
    Columns("U:U").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("U:U").Select
    Selection.Replace What:="-0.375", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     
End Sub

