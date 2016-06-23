Attribute VB_Name = "Module20"
Sub dimp5_series()

Dim msheet As Worksheet

Set msheet = ActiveSheet

    For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 3)
   Case Is = "TV"
    msheet.Cells(i, 3) = "Series"
   End Select

  Next
  MsgBox "done!"
  
End Sub

