Attribute VB_Name = "Module107"
Sub recurly_subs_box_order3()

Dim msheet As Worksheet
Dim n5 As String
Dim hm As Date

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Set msheet = ActiveSheet

n5 = "canceled"
hm = "24/03/2016"

With ActiveSheet
    For i = 2 To .UsedRange.Rows.Count
        If .Cells(i, 5) = n5 Then
            If .Cells(i, 36) < hm Then .Cells(i, 40) = "Currently No"
        If .Cells(i, 5) = n5 Then
            If .Cells(i, 36) >= hm Then .Cells(i, 40) = "Definitely Yes"
        End If
        End If
    Next i
End With

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
