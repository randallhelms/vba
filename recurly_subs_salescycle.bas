Attribute VB_Name = "Module70"
Sub recurly_subs_salescycle()

Dim msheet As Worksheet
Dim sc1 As Date
Dim sc2 As Date
Dim sc3 As Date
Dim sc4 As Date

Set msheet = ActiveSheet

sc1 = "06/10/2015"
sc2 = "24/11/2015"
sc3 = "01/12/2015"
sc4 = "24/01/2016"

Range("AM1").Select
ActiveCell = "sales_cycle"

For i = 2 To msheet.UsedRange.Rows.Count
  Select Case msheet.Cells(i, 33)
  Case Is >= sc4
  msheet.Cells(i, 39) = "4 - Hatsune Miku"
  GoTo ExitSelect
  Case Is >= sc3
  msheet.Cells(i, 39) = "3 - Naruto"
  GoTo ExitSelect
  Case Is >= sc2
  msheet.Cells(i, 39) = "2 - Black Friday"
  Case Is < sc2
  msheet.Cells(i, 39) = "1 - Kill La Kill"
  End Select

ExitSelect:
  Next i

End Sub
