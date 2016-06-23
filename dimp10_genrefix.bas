Attribute VB_Name = "Module37"
Sub dimp10_genrefix()

Dim msheet As Worksheet

Set msheet = ActiveSheet


  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 8)
   Case Is = "Action "
    msheet.Cells(i, 8) = "Action & Adventure"
    Case Is = "Action"
    msheet.Cells(i, 8) = "Action & Adventure"
    Case Is = "Adventure"
    msheet.Cells(i, 8) = "Action & Adventure"
   Case Is = "Thrillers"
    msheet.Cells(i, 8) = "Thriller"
   End Select
  Next
  MsgBox "done!"
End Sub
