Attribute VB_Name = "Module15"
Sub Raw_Data_Split()
Attribute Raw_Data_Split.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Raw_Data_Split Macro
'

'
    Columns("I:I").Select
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
    Columns("J:L").Select
    Selection.NumberFormat = "0"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Day"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("N1").Select
End Sub
