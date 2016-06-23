Attribute VB_Name = "Module73"
Sub rollup_delete_rows()
'
' rollup_delete_rows Macro
'

'
    Rows("1:6").Select
    Range("A6").Activate
    Selection.Delete Shift:=xlUp
End Sub

