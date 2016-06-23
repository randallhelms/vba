Attribute VB_Name = "Module116"
Sub localytics_region()

'set region for Localytics exports

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Range("D1").Select
ActiveCell = "Region"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "us"
    msheet.Cells(i, 4) = "1 - US"
   Case Is = "gb"
    msheet.Cells(i, 4) = "2 - UK & IE"
    Case Is = "uk"
    msheet.Cells(i, 4) = "2 - UK & IE"
    Case Is = "ie"
    msheet.Cells(i, 4) = "2 - UK & IE"
    Case Is = "at"
    msheet.Cells(i, 4) = "3 - DACH"
   Case Is = "ch"
    msheet.Cells(i, 4) = "3 - DACH"
    Case Is = "de"
    msheet.Cells(i, 4) = "3 - DACH"
    Case Is = "au"
    msheet.Cells(i, 4) = "5 - AU"
   Case Is = "nl"
    msheet.Cells(i, 4) = "6 - Benelux"
   Case Is = "be"
    msheet.Cells(i, 4) = "6 - Benelux"
   Case Is = "dk"
    msheet.Cells(i, 4) = "4 - Nordics"
   Case Is = "se"
    msheet.Cells(i, 4) = "4 - Nordics"
   Case Is = "no"
    msheet.Cells(i, 4) = "4 - Nordics"
   Case Is = "fi"
    msheet.Cells(i, 4) = "4 - Nordics"
   Case Is = "es"
    msheet.Cells(i, 4) = "7 - FR, IT & ES"
   Case Is = "fr"
    msheet.Cells(i, 4) = "7 - FR, IT & ES"
    Case Is = "it"
    msheet.Cells(i, 4) = "7 - FR, IT & ES"
    Case Else
    msheet.Cells(i, 4) = "8 - ROW"
   End Select
    
  Next i
  
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub


