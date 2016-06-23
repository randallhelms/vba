Attribute VB_Name = "Module93"
Sub recurly_subs_test_overwrite()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'Overwrites for test and reactivated cells

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 7)
   Case Is = "test"
    msheet.Cells(i, 5) = "test"
   Case Is = "reactivated"
    msheet.Cells(i, 5) = "reactivated"
   End Select
   
Next i
   
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
