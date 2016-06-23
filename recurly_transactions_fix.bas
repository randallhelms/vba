Attribute VB_Name = "Module101"
Sub recurly_transactions_fix()

Dim msheet As Worksheet
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

Set msheet = ActiveSheet

'Set purchases to test status

  For i = 2 To msheet.UsedRange.Rows.Count
   Select Case msheet.Cells(i, 2)
   Case Is = "Igor Sosedov"
    msheet.Cells(i, 4) = "test"
   Case Is = "Alex Sidorenko"
    msheet.Cells(i, 4) = "test"
    Case Is = "Randall Helms"
    msheet.Cells(i, 4) = "test"
    Case Is = "Tzvetan Horozov"
    msheet.Cells(i, 4) = "test"
    Case Is = "Gleb Bakhtin"
    msheet.Cells(i, 4) = "test"
   Case Is = "Rob Pereyda"
    msheet.Cells(i, 4) = "test"
    Case Is = "Harry Fuecks"
    msheet.Cells(i, 4) = "test"
   Case Is = "Harold Fuecks"
    msheet.Cells(i, 4) = "test"
   Case Is = "Robert Pereyda"
    msheet.Cells(i, 4) = "test"
   Case Is = "Simon Zseboek"
    msheet.Cells(i, 4) = "test"
   End Select
  
Next i

'Create separate date column

    Columns("H:H").Select
    Selection.Copy
    Range("AG1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AG1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("AH:AI").Select
    Selection.ClearContents

'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet

'Determine the data range you want to pivot
  SrcData = ActiveSheet.Name & "!" & Range("A1:AG10000").Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("amount"), "Sum of amount", xlSum
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("type")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("status")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("date2")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub

