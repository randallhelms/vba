Attribute VB_Name = "Module42"
Sub monthly_report_source()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("M7").Select
ActiveCell = "Group"

  For i = 8 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 1)
   Case Is = "redirect"
    msheet.Cells(i, 13) = "rd"
   Case Is = "(none)"
    msheet.Cells(i, 13) = "d"
    Case Is = "organic"
    msheet.Cells(i, 13) = "o"
   Case Is = "cpc"
    msheet.Cells(i, 13) = "a"
   Case Is = "banner"
    msheet.Cells(i, 13) = "a"
   Case Is = "CPC"
    msheet.Cells(i, 13) = "a"
   Case Else
    msheet.Cells(i, 13) = "r"
   End Select


  Next
  MsgBox "done!"
  
End Sub

