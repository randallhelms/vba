Attribute VB_Name = "Module71"
Sub rollup_naruto_dates()

Dim msheet As Worksheet
Dim sc1 As Integer
Dim sc2 As Integer
Dim sc3 As Integer

Set msheet = ActiveSheet

sc1 = "06/10/2015"
sc2 = "24/11/2015"
sc3 = "24/01/2016"

For i = 2 To msheet.UsedRange.Rows.Count
  Select Case msheet.Cells(i, 1)
  Case Is >= sc3
  msheet.Cells(i, 1) = "Sales Cycle 3"
  GoTo ExitSelect
  Case Is >= sc2
  msheet.Cells(i, 1) = "Sales Cycle 2"
  Case Is < sc2
  msheet.Cells(i, 36) = "Sales Cycle 1"
  End Select

ExitSelect:
  Next
  MsgBox "done!"

End Sub

sc1 = "06/10/2015 00:00:00 CET"
sc2 = "24/11/2015 09:00:00 CET"
sc3 = "01/12/2015 09:00:00 CET"
sc4 = "24/01/2016 09:00:00 CET"
