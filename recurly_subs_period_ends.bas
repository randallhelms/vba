Attribute VB_Name = "Module99"
Sub recurly_subs_period_ends()

'
' split period data
'
    Columns("M:M").Select
    Selection.Copy
    Range("AJ1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AJ1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "period_ends_date"
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "period_ends_time"
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "period_ends_timezone"
    
End Sub

