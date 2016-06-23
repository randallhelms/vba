Attribute VB_Name = "Module74"
Sub rollup_pivot()
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet
'Source: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

'Determine the data range you want to pivot
  SrcData = ActiveSheet.Name & "!" & Range("A1:H10000").Address(ReferenceStyle:=xlR1C1)

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
        "PivotTable1").PivotFields("Video Views"), "Sum of amount", xlSum
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("Country")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Movie ID")
        .Orientation = xlRowField
        .Position = 1
    End With

End Sub



