Attribute VB_Name = "Module112"
Sub dimp_match()


' copy existing data from title list

Dim wb1 As Window

Set wb1 = ActiveWindow


    Application.WindowState = xlNormal
    Windows("Title Info (Current).xlsx").Activate
    Columns("A:A").Select
    Selection.Copy
    wb1.Activate
    Range("M1").Select
    ActiveSheet.Paste

End Sub
