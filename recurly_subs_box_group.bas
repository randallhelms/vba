Attribute VB_Name = "Module59"
Sub recurly_subs_box_group()

Dim msheet As Worksheet

Set msheet = ActiveSheet

    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "box_group"
    
  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 33)
   Case Is < "24/11/2015"
   msheet.Cells(i, 36) = "Kill La Kill"
   Case Is > "02/12/2015"
   msheet.Cells(i, 36) = "Naruto"
    End Select
    
  Next
  MsgBox "done!"
End Sub

