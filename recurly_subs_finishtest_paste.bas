Attribute VB_Name = "Module105"
Sub recurly_subs_finishtest_paste()

Dim msheet As Worksheet
Dim raw As Worksheet
Dim wb1 As Window

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

Set msheet = ActiveSheet
Set raw = Worksheets("RawData")
Set wb1 = ActiveWindow

' Copy & Paste to RawData

    Columns("A:E").Select
    Selection.Copy
    Sheets("RawData").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
' Copy & Paste to Test Data

    msheet.Activate
    Columns("A:E").Select
    Selection.Copy
    Windows("test_list.xlsx").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
' Switch back to old Workbook

    wb1.Activate
    Sheets("RawData").Activate
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

