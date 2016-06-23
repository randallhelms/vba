Attribute VB_Name = "Module17"
Sub dimp2_epdist()
Attribute dimp2_epdist.VB_ProcData.VB_Invoke_Func = " \n14"
'
' dimp2_epdist Macro
'

Dim msheet As Worksheet

Set msheet = ActiveSheet

    Columns("E:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Distributor ID"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Episode ID"
    Range("G1").Select
    

  
End Sub

