Attribute VB_Name = "Module62"
Sub bens_report()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Range("E1").Select
ActiveCell = "Region"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "United States"
    msheet.Cells(i, 5) = "1 - US"
   Case Is = "United Kingdom"
    msheet.Cells(i, 5) = "2 - UK & IE"
    Case Is = "Switzerland"
    msheet.Cells(i, 5) = "3 - DACH"
    Case Is = "Sweden"
    msheet.Cells(i, 5) = "6 - Nordics"
    Case Is = "Denmark"
    msheet.Cells(i, 5) = "6 - Nordics"
    Case Is = "Finland"
    msheet.Cells(i, 5) = "6 - Nordics"
    Case Is = "Norway"
    msheet.Cells(i, 5) = "6 - Nordics"
    Case Is = "Netherlands"
    msheet.Cells(i, 5) = "4 - Benelux"
   Case Is = "Belgium"
    msheet.Cells(i, 5) = "4 - Benelux"
    Case Is = "Ireland"
    msheet.Cells(i, 5) = "2 - UK & IE"
    Case Is = "Germany"
    msheet.Cells(i, 5) = "3 - DACH"
   Case Is = "France"
    msheet.Cells(i, 5) = "5 - FR"
   Case Is = "Austria"
    msheet.Cells(i, 5) = "3 - DACH"
   Case Is = "Australia"
    msheet.Cells(i, 5) = "7 - AU"
    Case Else
    msheet.Cells(i, 5) = "8 - ROW"
   End Select

  Next i
  
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
