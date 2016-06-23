Attribute VB_Name = "Module100"
Sub recurly_subs_email_sort()

Dim msheet As Worksheet
Set msheet = ActiveSheet
Dim iCounter As Integer

Application.ScreenUpdating = False

'Sort by email field

   Range("c1") = "Index"
   Columns("A:k").Sort key1:=Range("c2"), _
      order1:=xlAscending, Header:=xlYes
   
Application.ScreenUpdating = True

End Sub
