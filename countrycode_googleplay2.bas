Attribute VB_Name = "Module111"
Sub countrycode_googleplay2()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("J1").Select
ActiveCell = "Tier"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 5)
   Case Is = "US"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "GB"
    msheet.Cells(i, 10) = "Tier 1"
    Case Is = "UK"
    msheet.Cells(i, 10) = "Tier 1"
    Case Is = "AT"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "CH"
    msheet.Cells(i, 10) = "Tier 1"
    Case Is = "DE"
    msheet.Cells(i, 10) = "Tier 1"
    Case Is = "AU"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "NL"
    msheet.Cells(i, 10) = "Tier 2"
   Case Is = "BE"
    msheet.Cells(i, 10) = "Tier 2"
   Case Is = "DK"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "SE"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "NO"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "FI"
    msheet.Cells(i, 10) = "Tier 1"
   Case Is = "ES"
    msheet.Cells(i, 10) = "Tier 2"
   Case Is = "FR"
    msheet.Cells(i, 10) = "Tier 2"
    Case Is = "IT"
    msheet.Cells(i, 10) = "Tier 2"
    Case Else
    msheet.Cells(i, 10) = "Tier 3"
   End Select
    
  Next
  MsgBox "done!"
End Sub



