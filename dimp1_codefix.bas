Attribute VB_Name = "Module16"
Sub dimp1_codefix()

Dim msheet As Worksheet

Set msheet = ActiveSheet


  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 3)
   Case Is = "Movie"
    msheet.Cells(i, 4) = msheet.Cells(i, 1)
    msheet.Cells(i, 5) = "<n/a>"
   End Select

  Next
  MsgBox "done!"
End Sub
