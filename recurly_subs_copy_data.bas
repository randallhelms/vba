Attribute VB_Name = "Module83"
Sub recurly_subs_copy_data()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'copy data from table and add totals

    Range("A3:A7").Select
    Selection.Copy
    Range("A11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D3:H7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D11:D15").Select
    Selection.Cut
    Range("B11").Select
    Selection.Insert Shift:=xlToRight
    Range("H11").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("H12").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6]:RC[-2])"
    Range("H12").Select
    Selection.AutoFill Destination:=Range("H12:H15"), Type:=xlFillDefault
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"
    Range("B17").Select
    Selection.AutoFill Destination:=Range("B17:H17"), Type:=xlFillDefault
    Range("B12:H17").Select
    Selection.NumberFormat = "#,##0"
    
' paste and clear

    Range("A11:F15").Select
    Selection.Copy
    Range("A19").Select
    ActiveSheet.Paste
    Range("A25").Select
    ActiveSheet.Paste
    Range("B20:F23").Select
    Selection.ClearContents
    Range("B26:F29").Select
    Selection.ClearContents
     
'first percentage group

    Range("B20").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[6]"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[5]"
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[4]"
    Range("E20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[3]"
    Range("F20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[2]"
    Range("B20:F20").Select
    Selection.AutoFill Destination:=Range("B20:F23"), Type:=xlFillDefault
   
'Create percentages by sales status

    Range("B26").Select
    ActiveCell.FormulaR1C1 = "=R[-14]C/R[-9]C"
    Range("B27").Select
    ActiveCell.FormulaR1C1 = "=R[-14]C/R[-10]C"
    Range("B28").Select
    ActiveCell.FormulaR1C1 = "=R[-14]C/R[-11]C"
    Range("B29").Select
    ActiveCell.FormulaR1C1 = "=R[-14]C/R[-12]C"
    Range("B26:B29").Select
    Selection.AutoFill Destination:=Range("B26:F29"), Type:=xlFillDefault
    
'make all percentages

    Range("B20:F29").Select
    Selection.Style = "Percent"

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

