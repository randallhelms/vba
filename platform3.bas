Attribute VB_Name = "Module31"
Sub platform3()

Dim msheet As Worksheet

Set msheet = ActiveSheet


  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 7)
   Case Is = "MBA"
    msheet.Cells(i, 9) = "Apple"
   Case Is = "MBL"
    msheet.Cells(i, 9) = "Android"
       Case Is = "MWS"
    msheet.Cells(i, 9) = "Web"
   Case Is = "PC"
    msheet.Cells(i, 9) = "Web"
   Case Is = "STB"
    msheet.Cells(i, 9) = "CTV"
   Case Is = "TBA"
    msheet.Cells(i, 9) = "Apple"
   Case Is = "TBL"
    msheet.Cells(i, 9) = "Android"
    Case Is = "TVI"
    msheet.Cells(i, 9) = "CTV"
   End Select
  Next
  MsgBox "done!"
End Sub

