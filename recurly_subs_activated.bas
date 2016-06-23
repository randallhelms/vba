Attribute VB_Name = "Module63"
Sub recurly_subs_activated()
Attribute recurly_subs_activated.VB_ProcData.VB_Invoke_Func = " \n14"
'
' recurly_subs_activated Macro
'

    Columns("R:R").Select
    Selection.Copy
    Range("AG1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AG1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "activated_at_date"
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "activated_at_time"
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "activated_at_timezone"
    Range("AG2").Select
End Sub
