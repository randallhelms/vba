Attribute VB_Name = "Module57"
Sub recurly_transactions_test()

Dim msheet As Worksheet

Set msheet = ActiveSheet

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "Igor Sosedov"
    msheet.Cells(i, 4) = "test"
   Case Is = "Alex Sidorenko"
    msheet.Cells(i, 4) = "test"
    Case Is = "Randall Helms"
    msheet.Cells(i, 4) = "test"
    Case Is = "Tzvetan Horozov"
    msheet.Cells(i, 4) = "test"
    Case Is = "Gleb Bakhtin"
    msheet.Cells(i, 4) = "test"
   Case Is = "Rob Pereyda"
    msheet.Cells(i, 4) = "test"
    Case Is = "Harry Fuecks"
    msheet.Cells(i, 4) = "test"
   End Select
    
  Next
  MsgBox "done!"
End Sub



