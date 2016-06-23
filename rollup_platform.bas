Attribute VB_Name = "Module72"
Sub rollup_platform()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("H7").Select
ActiveCell = "Platform"

  For i = 8 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 3)
   Case Is = "UA-24364238-23"
    msheet.Cells(i, 9) = "Android"
   Case Is = "UA-24364238-24"
    msheet.Cells(i, 9) = "iOS"
    Case Is = "UA-24364238-2"
    msheet.Cells(i, 9) = "CTV"
    Case Is = "UA-24364238-38"
    msheet.Cells(i, 9) = "Web"
    Case Is = "UA-24364238-26"
    msheet.Cells(i, 9) = "Roku"
    Case Is = "UA-24364238-43"
    msheet.Cells(i, 9) = "Flipps"
    End Select
    
  Next
  MsgBox "done!"
End Sub



