Attribute VB_Name = "Module87"
Sub recurly_subs_remove_timezone()
    
'removes timezone
    
    Columns("L:U").Select
    Selection.Replace What:=" CET", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" CEST", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub
