Attribute VB_Name = "Module50"
Sub dimp12_branded_channels()

Dim msheet As Worksheet

Set msheet = ActiveSheet

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 12)
   Case Is = "1324"
    msheet.Cells(i, 8) = "Branded Channel"
   Case Is = "1322"
    msheet.Cells(i, 8) = "Branded Channel"
   Case Is = "1316"
    msheet.Cells(i, 8) = "Branded Channel"
   Case Is = "1317"
    msheet.Cells(i, 8) = "Branded Channel"
   Case Is = "1319"
    msheet.Cells(i, 8) = "Branded Channel"
   End Select
    
  Next
  MsgBox "done!"
End Sub


