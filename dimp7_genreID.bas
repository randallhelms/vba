Attribute VB_Name = "Module22"
Sub dimp7_genreID()

Dim msheet As Worksheet

Set msheet = ActiveSheet

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 8)
   Case Is = "Glamour"
    msheet.Cells(i, 18) = "0"
   Case Is = "New Releases"
    msheet.Cells(i, 18) = "1"
   Case Is = "Action"
    msheet.Cells(i, 18) = "2"
   Case Is = "Action & Adventure"
    msheet.Cells(i, 18) = "2"
   Case Is = "Animation"
    msheet.Cells(i, 18) = "3"
   Case Is = "Kids"
    msheet.Cells(i, 18) = "5"
   Case Is = "Classics"
    msheet.Cells(i, 18) = "6"
    Case Is = "Comedy"
    msheet.Cells(i, 18) = "7"
   Case Is = "Crime & Gangsters"
    msheet.Cells(i, 18) = "8"
   Case Is = "Documentary"
    msheet.Cells(i, 18) = "9"
   Case Is = "Drama"
    msheet.Cells(i, 18) = "10"
   Case Is = "Horror"
    msheet.Cells(i, 18) = "11"
   Case Is = "Festival Darlings"
    msheet.Cells(i, 18) = "12"
   Case Is = "Music & Musicals"
    msheet.Cells(i, 18) = "13"
   Case Is = "Romance"
    msheet.Cells(i, 18) = "14"
    Case Is = "Sci-Fi & Fantasy"
    msheet.Cells(i, 18) = "15"
      Case Is = "Sports"
    msheet.Cells(i, 18) = "18"
   Case Is = "Thrillers"
    msheet.Cells(i, 18) = "20"
   Case Is = "Korean Drama"
    msheet.Cells(i, 18) = "50"
   Case Is = "Anime"
    msheet.Cells(i, 18) = "58"
   Case Is = "Viewster Fest"
    msheet.Cells(i, 18) = "59"
   
    End Select
    
  Next
  MsgBox "done!"
  
End Sub

