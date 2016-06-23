Attribute VB_Name = "Module102"
Sub recurly_subs_expired_tests()

With ActiveSheet
    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 5) = "expired" Then
            If Application.IsNA(.Cells(i, 6)) Then .Cells(i, 7) = "test"
        End If
    Next i
End With

End Sub

