Attribute VB_Name = "Module55"
Sub recurly_transactions_date()
Attribute recurly_transactions_date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' recurly_transactions_date Macro
'
    Columns("H:H").Select
    Selection.Copy
    Range("AK1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AK1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("AL:AM").Select
    Selection.ClearContents

End Sub
