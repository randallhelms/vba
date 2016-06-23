Attribute VB_Name = "Module76"
Sub rollup_platform2()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Range("j1").Select
ActiveCell = "Platform"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 3)
   Case Is = "UA-24364238-23"
    msheet.Cells(i, 10) = "Android"
   Case Is = "UA-24364238-24"
    msheet.Cells(i, 10) = "iOS"
    Case Is = "UA-24364238-2"
    msheet.Cells(i, 10) = "CTV"
    Case Is = "UA-24364238-38"
    msheet.Cells(i, 10) = "Web"
    Case Is = "UA-24364238-26"
    msheet.Cells(i, 10) = "Roku"
    Case Is = "UA-24364238-43"
    msheet.Cells(i, 10) = "Flipps"
    End Select
    
  Next
  MsgBox "done!"
End Sub

