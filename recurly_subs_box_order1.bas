Attribute VB_Name = "Module97"
Sub recurly_subs_box_order1()

Dim msheet As Worksheet
Dim n1 As String
Dim n2 As String
Dim n3 As String
Dim hm As Date

Application.ScreenUpdating = False

Set msheet = ActiveSheet

n1 = "test"
n2 = "reactivated"
n3 = "expired"
hm = "24/03/2016"

For i = 2 To msheet.UsedRange.Rows.Count
  Select Case msheet.Cells(i, 5)
  Case Is = n1
  msheet.Cells(i, 40) = "n/a"
  Case Is = n2
  msheet.Cells(i, 40) = "Reactivated"
  Case Is = n3
  msheet.Cells(i, 40) = "Definitely Not"
  End Select

Next i

Application.ScreenUpdating = True

End Sub
