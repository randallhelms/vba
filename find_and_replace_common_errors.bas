Attribute VB_Name = "Module32"
Sub find_and_replace_common_errors()
Attribute find_and_replace_common_errors.VB_ProcData.VB_Invoke_Func = " \n14"
'
' find_and_replace_common_errors Macro
'

'
    Columns("D:F").Select
    Selection.Replace What:="�", Replacement:=" - ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:=" - ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="o", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:=" ...", Replacement:=" ...", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="a", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="i", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="e", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="pu", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="'", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    
    
End Sub
