Attribute VB_Name = "Module64"
Sub countrycode_localytics()

'set region for Localytics exports

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Range("G1").Select
ActiveCell = "Region"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 4)
   Case Is = "us"
    msheet.Cells(i, 7) = "1 - US"
   Case Is = "gb"
    msheet.Cells(i, 7) = "2 - UK & IE"
    Case Is = "uk"
    msheet.Cells(i, 7) = "2 - UK & IE"
    Case Is = "ie"
    msheet.Cells(i, 7) = "2 - UK & IE"
    Case Is = "at"
    msheet.Cells(i, 7) = "3 - DACH"
   Case Is = "ch"
    msheet.Cells(i, 7) = "3 - DACH"
    Case Is = "de"
    msheet.Cells(i, 7) = "3 - DACH"
    Case Is = "au"
    msheet.Cells(i, 7) = "5 - AU"
   Case Is = "nl"
    msheet.Cells(i, 7) = "6 - Benelux"
   Case Is = "be"
    msheet.Cells(i, 7) = "6 - Benelux"
   Case Is = "dk"
    msheet.Cells(i, 7) = "4 - Nordics"
   Case Is = "se"
    msheet.Cells(i, 7) = "4 - Nordics"
   Case Is = "no"
    msheet.Cells(i, 7) = "4 - Nordics"
   Case Is = "fi"
    msheet.Cells(i, 7) = "4 - Nordics"
   Case Is = "es"
    msheet.Cells(i, 7) = "7 - FR, IT & ES"
   Case Is = "fr"
    msheet.Cells(i, 7) = "7 - FR, IT & ES"
    Case Is = "it"
    msheet.Cells(i, 7) = "7 - FR, IT & ES"
    Case Else
    msheet.Cells(i, 7) = "8 - ROW"
   End Select
    
  Next i
  
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

