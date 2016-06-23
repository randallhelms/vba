Attribute VB_Name = "Module117"
Sub localytics_title_info()

Dim msheet As Worksheet
Dim lastRow As Long
Dim mwindow As Window

Set msheet = ActiveSheet
Set mwindow = ActiveWindow
lastRow = Range("A" & Rows.Count).End(xlUp).Row

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

' Add column titles

    Range("E1").Select
    ActiveCell = "Content Type"
    Range("F1").Select
    ActiveCell = "Content Title"
    Range("G1").Select
    ActiveCell = "Episode Number"
    Range("H1").Select
    ActiveCell = "Genre"
    Range("I1").Select
    ActiveCell = "Content Owner"
    
' Add formulas

    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-4],R2C11:R25000C11,R2C19:R25000C19)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-5],R2C11:R25000C11,R2C14:R25000C14)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-6],R2C11:R25000C11,R2C16:R25000C16)"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-7],R2C11:R25000C11,R2C17:R25000C17)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=LOOKUP(RC[-8],R2C11:R25000C11,R2C12:R25000C12)"

' Drag formulas to lastRow

    Range("E2:I2").Select
    Range("E2:I2").AutoFill Destination:=Range("E2:I" & lastRow)

' Copy and paste Title Info

    Windows("title_info.xlsx").Activate
    Columns("A:K").Select
    Selection.Copy
    mwindow.Activate
    Range("K1").Select
    ActiveSheet.Paste
    
' Remove spaces from Movie IDs (if they haven't been removed previously)
    
    Range("K:K").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

' Paste as values in columns E-I

    Columns("E:I").Select
    Selection.Copy
    Range("E1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

' Remove Title Info

    Columns("K:V").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
