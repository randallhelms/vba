Attribute VB_Name = "Module35"
Sub dimp9_replacements()

Dim msheet As Worksheet

Set msheet = ActiveSheet

   Columns("I:K").Select
    Selection.Replace What:=",", Replacement:=", ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="", Replacement:="<n/a>", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

 
End Sub
