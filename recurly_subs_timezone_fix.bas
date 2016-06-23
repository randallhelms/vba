Attribute VB_Name = "Module85"
Sub recurly_subs_timezone_fix()

Dim msheet As Worksheet
Dim sc1 As Date
Dim sc2 As Date
Dim sc3 As Date
Dim sc4 As Date
Dim nAM As String

Set msheet = ActiveSheet

sc1 = "06/10/2015"
sc2 = "24/11/2015"
sc3 = "01/12/2015"
sc4 = "24/01/2016"
nAM = "09:00:00"

For i = 2 To msheet.UsedRange.Rows.Count
If ((msheet.Cells(i, 33) = sc2) And (msheet.Cells(i, 34) < nAM)) Then
msheet.Cells(1, 36) = "1 - Kill La Kill"
End If

End Sub
