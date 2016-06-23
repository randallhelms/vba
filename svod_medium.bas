Attribute VB_Name = "Module49"
Sub svod_medium()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("D7").Select
ActiveCell = "Traffic Type"

  For i = 8 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 1)
   Case Is = "email"
    msheet.Cells(i, 4) = "Email"
   Case Is = "organic"
    msheet.Cells(i, 4) = "Organic Search"
    Case Is = "splash"
    msheet.Cells(i, 4) = "Splash"
    Case Is = "cpc"
    msheet.Cells(i, 4) = "Marketing"
    Case Is = "advertorial"
    msheet.Cells(i, 4) = "Marketing"
   Case Is = "takeover"
    msheet.Cells(i, 4) = "Marketing"
    Case Is = "banner"
    msheet.Cells(i, 4) = "Marketing"
    Case Is = "mobile"
    msheet.Cells(i, 4) = "Mobile"
    Case Is = "(none)"
    msheet.Cells(i, 4) = "Direct"
    Case Else
    msheet.Cells(i, 4) = "Referral"
   End Select
    
  Next
  MsgBox "done!"
End Sub



