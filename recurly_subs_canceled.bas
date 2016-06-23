Attribute VB_Name = "Module58"
Sub recurly_subs_canceled()
Attribute recurly_subs_canceled.VB_ProcData.VB_Invoke_Func = " \n14"
'
' split period data

    Columns("T:T").Select
    Selection.Copy
    Range("AA1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AA1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "canceled_at_date"
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "canceled_at_time"
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "canceled_at_timezone"

End Sub
