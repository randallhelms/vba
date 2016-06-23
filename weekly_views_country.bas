Attribute VB_Name = "Module44"
Sub weekly_views_country()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("E1").Select
ActiveCell = "Region"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 4)
   Case Is = "US"
    msheet.Cells(i, 5) = "US"
       Case Is = "us"
    msheet.Cells(i, 5) = "US"
   Case Is = "GB"
    msheet.Cells(i, 5) = "UK"
   Case Is = "gb"
    msheet.Cells(i, 5) = "UK"
       Case Is = "ca"
    msheet.Cells(i, 5) = "CA"
   Case Is = "CA"
    msheet.Cells(i, 5) = "CA"
    Case Else
    msheet.Cells(i, 5) = "ROW"
    End Select

  Next
  MsgBox "done!"
  
  End Sub
