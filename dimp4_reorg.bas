Attribute VB_Name = "Module19"
Sub dimp4_reorg()
'
' dimp4_reorg Macro
'

'
    Columns("D:D").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:O").Select
    Range("O1").Activate
    Selection.Cut
    Columns("N:P").Select
    Range("P1").Activate
    Application.CutCopyMode = False
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Genre ID"
  
    MsgBox "done!"
End Sub

