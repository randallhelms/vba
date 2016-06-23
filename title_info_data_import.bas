Attribute VB_Name = "Module115"
Sub title_info_data_import()

Dim msheet As Worksheet

Set msheet = ActiveSheet

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

'Change Col D to 'Movie ID', fill in Movie ID for movies, and set n/a for Episode Title

Range("D1").Select
ActiveCell.FormulaR1C1 = "Movie ID"
    
  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 3)
   Case Is = "Movie"
    msheet.Cells(i, 4) = msheet.Cells(i, 1)
    msheet.Cells(i, 5) = "<n/a>"
   End Select

Next i

' Delete unneeded columns

    Range("A:A,G:I,K:K,N:O").Select
    Selection.Delete Shift:=xlToLeft
    
' Create Distributor ID and Episode Number

    Columns("D:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Distributor ID"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Product ID"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Episode ID"

    Range("E:E").Select
    Selection.Delete Shift:=xlToLeft
    
' Set Episode ID for Movies to N/A

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "Movie"
    msheet.Cells(i, 5) = "<n/a>"
   End Select
   
   Next i
    
' Change TV to Series

    For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "TV"
    msheet.Cells(i, 2) = "Series"
   End Select
   
   Next i
    
' Set Branded Channels as separate genre

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 4)
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
   
   Next i
   
' Make the genres more consistent

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 8)
   Case Is = "Action "
    msheet.Cells(i, 8) = "Action & Adventure"
    Case Is = "Action"
    msheet.Cells(i, 8) = "Action & Adventure"
    Case Is = "Adventure"
    msheet.Cells(i, 8) = "Action & Adventure"
   Case Is = "Thrillers"
    msheet.Cells(i, 8) = "Thriller"
   End Select
  Next i
  
' Create a separate Genre ID

Range("K1").Select
ActiveCell.FormulaR1C1 = "Genre ID"

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 8)
   Case Is = "Glamour"
    msheet.Cells(i, 11) = "0"
   Case Is = "New Releases"
    msheet.Cells(i, 11) = "1"
   Case Is = "Action"
    msheet.Cells(i, 11) = "2"
   Case Is = "Action & Adventure"
    msheet.Cells(i, 11) = "2"
   Case Is = "Animation"
    msheet.Cells(i, 11) = "3"
   Case Is = "Kids"
    msheet.Cells(i, 11) = "5"
   Case Is = "Classics"
    msheet.Cells(i, 11) = "6"
    Case Is = "Comedy"
    msheet.Cells(i, 11) = "7"
   Case Is = "Crime & Gangsters"
    msheet.Cells(i, 11) = "8"
   Case Is = "Documentary"
    msheet.Cells(i, 11) = "9"
   Case Is = "Drama"
    msheet.Cells(i, 11) = "10"
   Case Is = "Horror"
    msheet.Cells(i, 11) = "11"
   Case Is = "Festival Darlings"
    msheet.Cells(i, 11) = "12"
   Case Is = "Music & Musicals"
    msheet.Cells(i, 11) = "13"
   Case Is = "Romance"
    msheet.Cells(i, 11) = "14"
    Case Is = "Sci-Fi & Fantasy"
    msheet.Cells(i, 11) = "15"
      Case Is = "Sports"
    msheet.Cells(i, 11) = "18"
   Case Is = "Thrillers"
    msheet.Cells(i, 11) = "20"
   Case Is = "Korean Drama"
    msheet.Cells(i, 11) = "50"
   Case Is = "Anime"
    msheet.Cells(i, 11) = "58"
   Case Is = "Viewster Fest"
    msheet.Cells(i, 11) = "59"
   
    End Select
    
    Next i
    
' Fix standard errors

    Range("A:A,F:H").Select
    Selection.Replace What:="—", Replacement:=" - ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="–", Replacement:=" - ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="‘", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ó", Replacement:="o", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:=" ...", Replacement:=" ...", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ã", Replacement:="a", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="í", Replacement:="i", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="É", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="µ", Replacement:="Mu", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="’", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ö", Replacement:="oe", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="sâ€”", Replacement:=" - ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â€”", Replacement:=": ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â€™", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â˜†", Replacement:="-", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="a‰", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ß", Replacement:="ss", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ç", Replacement:="c", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="§", Replacement:="ss", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="…", Replacement:=" ...", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â˜", Replacement:=":", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:=" ___ ", Replacement:="___", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="³", Replacement:="(Cubed)", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â€", Replacement:=":", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="²", Replacement:="(Squared)", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ê", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="â", Replacement:="a", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="è", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="é", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ë", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ï", Replacement:="i", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ÿ", Replacement:="y", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="à", Replacement:="a", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ù", Replacement:="u", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="î", Replacement:="i", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ô", Replacement:="o", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="û", Replacement:="u", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ñ", Replacement:="n", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="¦", Replacement:="!", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ü", Replacement:="ue", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ä", Replacement:="ae", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="á", Replacement:="a", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="´", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="ã", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="Ò", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="œ", Replacement:="u", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="Ê", Replacement:=" ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="¡", Replacement:=":", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="Ó", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="¯", Replacement:="O", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="À", Replacement:="?", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
' Make some improvements to the languages

    Range("I:J").Select
    Selection.Replace What:=",", Replacement:=", ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="", Replacement:="<n/a>", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

' Reorganize fields to match the data import format

    Columns("C:C").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:H").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    
' Remove spaces from Movie IDs
    
    Range("A:A").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
  
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
