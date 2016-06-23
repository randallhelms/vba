Attribute VB_Name = "Module114"
Sub aws_platform()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Range("F1").Select
ActiveCell = "Product"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 5)
   Case Is = "tbl"
    msheet.Cells(i, 6) = "Android"
   Case Is = "mbl"
    msheet.Cells(i, 6) = "Android"
   Case Is = "tba"
    msheet.Cells(i, 6) = "iOS"
   Case Is = "mba"
    msheet.Cells(i, 6) = "iOS"
    Case Is = "tvi"
    msheet.Cells(i, 6) = "CTV"
    Case Is = "stb"
    msheet.Cells(i, 6) = "CTV"
    Case Is = "pc"
    msheet.Cells(i, 6) = "Web"
    Case Is = "mws"
    msheet.Cells(i, 6) = "Web"
    End Select
    
    Next i
    
  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 4)
    Case Is = "roku"
    msheet.Cells(i, 6) = "Roku"
    End Select
    
  Next i

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
