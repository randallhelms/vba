Attribute VB_Name = "Module104"
Sub recurly_subs_order_sort()

Dim msheet As Worksheet
Set msheet = ActiveSheet

'Sort by order field

   Range("k1") = "Index"
   Columns("A:k").Sort key1:=Range("k2"), _
      order1:=xlAscending, Header:=xlYes

End Sub


