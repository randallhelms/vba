Attribute VB_Name = "Module106"
Sub recurly_subs_box_order2()

Dim msheet As Worksheet
Dim n4 As String
Dim hm As Date

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Set msheet = ActiveSheet

n4 = "active"
hm = "24/03/2016"

With ActiveSheet
    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 5) = n4 Then
            If .Cells(i, 36) < hm Then .Cells(i, 40) = "Currently Yes"
        If .Cells(i, 5) = n4 Then
            If .Cells(i, 36) >= hm Then .Cells(i, 40) = "Definitely Yes"
        End If
        End If
    Next i
End With

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

