Attribute VB_Name = "Module45"
Sub countrycode_russia()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("g1").Select
ActiveCell = "Russia"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 6)
   Case Is = "ru"
    msheet.Cells(i, 7) = "Russia"
   Case Is = "RU"
    msheet.Cells(i, 7) = "Russia"
   Case Else
    msheet.Cells(i, 7) = "ROW"
   End Select
    
  Next
  MsgBox "done!"
End Sub



