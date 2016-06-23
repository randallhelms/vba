Attribute VB_Name = "Module18"
Sub dimp3_epid()

Dim msheet As Worksheet

Set msheet = ActiveSheet


  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 7)
   Case Is = 0
    msheet.Cells(i, 7) = "<n/a>"
   End Select
   
  Next
  MsgBox "done!"
  
End Sub
