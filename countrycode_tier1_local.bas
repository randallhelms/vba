Attribute VB_Name = "Module108"
Sub countrycode_tier1_local()

'set region for Localytics exports

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Range("I1").Select
ActiveCell = "Region"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 5)
   Case Is = "us"
    msheet.Cells(i, 9) = "1 - US"
   Case Is = "gb"
    msheet.Cells(i, 9) = "2 - UK & IE"
    Case Is = "uk"
    msheet.Cells(i, 9) = "2 - UK & IE"
    Case Is = "at"
    msheet.Cells(i, 9) = "3 - DACH"
   Case Is = "ch"
    msheet.Cells(i, 9) = "3 - DACH"
    Case Is = "de"
    msheet.Cells(i, 9) = "3 - DACH"
    Case Else
    msheet.Cells(i, 9) = "8 - ROW"
   End Select
    
  Next i
  
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub


