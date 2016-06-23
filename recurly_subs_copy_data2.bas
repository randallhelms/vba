Attribute VB_Name = "Module92"
Sub recurly_subs_copy_data2()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'copy data from table and add totals

    Range("a2:e6").Select
    Selection.Copy
    Range("A11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G11").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("G12").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-5]:RC[-2])"
    Range("G12").Select
    Selection.AutoFill Destination:=Range("G12:G15"), Type:=xlFillDefault
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"
    Range("B17").Select
    Selection.AutoFill Destination:=Range("B17:G17"), Type:=xlFillDefault
    Range("B12:G17").Select
    Selection.NumberFormat = "#,##0"
  
' paste and clear

    Range("A11:E15").Select
    Selection.Copy
    Range("A19").Select
    ActiveSheet.Paste
    Range("B20:F23").Select
    Selection.ClearContents
    
'first percentage group

    Range("B20").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[5]"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[4]"
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[3]"
    Range("E20").Select
    ActiveCell.FormulaR1C1 = "=R[-8]C/R[-8]C[2]"
    Range("B20:E20").Select
    Selection.Style = "Percent"
    Selection.AutoFill Destination:=Range("B20:E23"), Type:=xlFillDefault
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub



