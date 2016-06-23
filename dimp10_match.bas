Attribute VB_Name = "Module46"
Sub dimp10_match()
Attribute dimp10_match.VB_ProcData.VB_Invoke_Func = " \n14"
'
' dimp10_match Macro
'

'
    Range("B:B,E:E").Select
    Range("E1").Activate
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Match"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MATCH(RC[-1],R2C5:R24308C5,0)"
    Range("C3").Select
End Sub
