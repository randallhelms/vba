Attribute VB_Name = "Module82"
Sub recurly_subs_pivot2()
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet
'Source: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

'Determine the data range you want to pivot
  SrcData = ActiveSheet.Name & "!" & Range("A1:AO13000").Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("uuid"), "Count of uuid", xlCount
        With ActiveSheet.PivotTables("PivotTable1").PivotFields("state")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("sales_cycle")
        .Orientation = xlRowField
        .Position = 1
    End With
    
ActiveSheet.Name = "CurrentStatus"

End Sub




