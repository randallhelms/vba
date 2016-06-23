Attribute VB_Name = "Module94"
Sub recurly_subs_test_next()
'
' copy test data from test list
'
Dim wb1 As Window

Set wb1 = ActiveWindow


    Application.WindowState = xlNormal
    Windows("test_list.xlsx").Activate
    Columns("A:G").Select
    Selection.Copy
    wb1.Activate
    Range("M1").Select
    ActiveSheet.Paste


End Sub

