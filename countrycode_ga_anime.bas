Attribute VB_Name = "Module68"
Sub countrycode_ga_anime()
Attribute countrycode_ga_anime.VB_ProcData.VB_Invoke_Func = "t\n14"

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("G7").Select
ActiveCell = "Order"

  For i = 8 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 1)
   Case Is = "United States"
    msheet.Cells(i, 7) = "12"
   Case Is = "United Kingdom"
    msheet.Cells(i, 7) = "11"
    Case Is = "Switzerland"
    msheet.Cells(i, 7) = "10"
    Case Is = "New Zealand"
    msheet.Cells(i, 7) = "9"
    Case Is = "Netherlands"
    msheet.Cells(i, 7) = "8"
   Case Is = "Mexico"
    msheet.Cells(i, 7) = "7"
    Case Is = "Ireland"
    msheet.Cells(i, 7) = "6"
    Case Is = "Germany"
    msheet.Cells(i, 7) = "5"
   Case Is = "France"
    msheet.Cells(i, 7) = "4"
   Case Is = "Canada"
    msheet.Cells(i, 7) = "3"
   Case Is = "Austria"
    msheet.Cells(i, 7) = "2"
   Case Is = "Australia"
    msheet.Cells(i, 7) = "1"
   End Select

Range("A8:g500").Sort _
key1:=Range("G8"), order1:=xlAscending

  Next
  MsgBox "done!"
End Sub



