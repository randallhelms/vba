Attribute VB_Name = "Module96"
Sub recurly_subs_legit()

Dim msheet As Worksheet
Dim n1 As String

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Set msheet = ActiveSheet

n1 = "test"

For i = 2 To msheet.UsedRange.Rows.Count
  Select Case msheet.Cells(i, 5)
  Case Is = n1
  msheet.Cells(i, 41) = "No"
  Case Else
  msheet.Cells(i, 41) = "Yes"
  End Select

ExitSelect:
  Next i

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
