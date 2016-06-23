Attribute VB_Name = "Module77"
Sub recurly_subs_test_start()

Dim msheet As Worksheet
Dim lastRow As Long
lastRow = Range("A" & Rows.Count).End(xlUp).Row

' Paste from Previous Section

    Columns("A:E").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste

' Add formulas

    Range("F1").Select
    ActiveCell = "Match"
    Range("G1").Select
    ActiveCell = "UID Status"
    Range("H1").Select
    ActiveCell = "Mismatch"
    Range("I1").Select
    ActiveCell = "Dupe"
    Range("J1").Select
    ActiveCell = "Email Match"
    Range("f2").Select
    ActiveCell.FormulaR1C1 = "=MATCH(RC[-5],R2C13:R13000C13,0)"
    Range("g2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-6],R2C13:R13000C13,R2C17:R13000C17)"
    Range("h2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=RC[-3],0,1)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[1]C[-8]=RC[-8],0,1)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IF(R[1]C[-7]=RC[-7],0,1)"
    Range("F2:J2").Select
    Range("F2:J2").AutoFill Destination:=Range("F2:J" & lastRow)
    
' Add Order Section

    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Order"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("k2:k4").Select
    Selection.AutoFill Destination:=Range("k2:k" & lastRow)

' Add Filter and Freeze

    Range("A1:K1").Select
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
End Sub

