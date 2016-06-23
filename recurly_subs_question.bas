Attribute VB_Name = "Module78"
Sub recurly_subs_question()

Dim msheet As Worksheet

    Range("AN1").Select
    ActiveCell = "Will they get next box?"
    Range("AO1").Select
    ActiveCell = "Legit?"
    Selection.AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
End Sub
