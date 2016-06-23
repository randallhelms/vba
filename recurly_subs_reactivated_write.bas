Attribute VB_Name = "Module103"
Sub recurly_subs_reactivated_write()

'Overwrites for reactivated accounts
Application.ScreenUpdating = False
Application.DisplayStatusBar = False

With ActiveSheet
    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 10) = "0" Then
            If .Cells(i, 5) = "expired" Or .Cells(i, 5) = "cancelled" Then
            .Cells(i, 7) = "reactivated"
        End If
        End If
    Next i
End With

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

