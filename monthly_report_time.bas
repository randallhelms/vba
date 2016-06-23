Attribute VB_Name = "Module41"
Sub monthly_report_time()
Attribute monthly_report_time.VB_ProcData.VB_Invoke_Func = " \n14"
'
' monthly_report_time Macro
'
Dim lastRow As Integer
lastRow = ActiveSheet.UsedRange.Rows.Count

    Range("L7").Select
    ActiveCell = "Total Time"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]*RC[-7]"
    Range("L8").Select
    Selection.AutoFill Destination:=Range("L8:L50")
    Range("L8:L50").Select
    Columns("L:L").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight

 
End Sub

