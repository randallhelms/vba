Attribute VB_Name = "Module12"
Sub countrycode()

Dim msheet As Worksheet

Set msheet = ActiveSheet


  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 7)
   Case Is = "us"
    msheet.Cells(i, 10) = "1 - US"
   Case Is = "gb"
    msheet.Cells(i, 10) = "2 - UK & IE"
    Case Is = "uk"
    msheet.Cells(i, 10) = "2 - UK & IE"
    Case Is = "ie"
    msheet.Cells(i, 10) = "2 - UK & IE"
    Case Is = "at"
    msheet.Cells(i, 10) = "3 - DACH"
   Case Is = "ch"
    msheet.Cells(i, 10) = "3 - DACH"
    Case Is = "de"
    msheet.Cells(i, 10) = "3 - DACH"
    Case Is = "au"
    msheet.Cells(i, 10) = "4 - AU"
   Case Is = "nl"
    msheet.Cells(i, 10) = "5 - Benelux"
   Case Is = "be"
    msheet.Cells(i, 10) = "5 - Benelux"
   Case Is = "dk"
    msheet.Cells(i, 10) = "6 - Nordic"
   Case Is = "se"
    msheet.Cells(i, 10) = "6 - Nordic"
   Case Is = "no"
    msheet.Cells(i, 10) = "6 - Nordic"
   Case Is = "fi"
    msheet.Cells(i, 10) = "6 - Nordic"
   Case Is = "es"
    msheet.Cells(i, 10) = "7 - ES"
   Case Is = "fr"
    msheet.Cells(i, 10) = "8 - FR"
    Case Else
    msheet.Cells(i, 10) = "9 - ROW"
   End Select
    
  Next
  MsgBox "done!"
End Sub

