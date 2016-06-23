Attribute VB_Name = "Module60"
Sub recurly_subs_expires_date()
Attribute recurly_subs_expires_date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' recurly_subs_expires_date Macro
'

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

    Columns("U:U").Select
    Selection.Copy
    Range("AD1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "expires_date"
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "expires_time"
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "expires_time_zone"
    Range("AD2").Select

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

