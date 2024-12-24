Option Explicit
#Const cDebugMode = 0
Sub Main()
    Dim cltSales As Collection, sSheetName As String, sDataSheetName As String, sAnalysisSheetName As String, sTableName As String, RawDataName As Variant
    Dim arrDefaultHeaders As Variant, arrRawData As Variant
    sSheetName = "database"
    sDataSheetName = "table"
    sAnalysisSheetName = "analysis"
    sTableName = "tblCleanData"
    arrRawData = Array("sales_21_22", "sales_22_23")
    arrDefaultHeaders = Array("Year", "Month", "Store Type", "Variable", "Value")
    Set cltSales = New Collection
    
    ' Main code
    Application.ScreenUpdating = False
    For Each RawDataName In arrRawData
        Call UpdateData.GetData(sSheetName, CStr(RawDataName), cltSales)
    Next
    Call UpdateData.WriteTable(sDataSheetName, sTableName, cltSales, arrDefaultHeaders)
    Call UpdateData.UpdatePivotTable(sDataSheetName, sAnalysisSheetName, sTableName, arrDefaultHeaders)
    Application.ScreenUpdating = True
    
    ' Debugging, ignore it
    #If cDebugMode = 1 Then
        Dim oItem As clsMonthlySales
        For Each oItem In cltSales
            Debug.Print oItem.MonthYear
        Next
    #End If
    Set cltSales = Nothing
    
End Sub
Sub DefinePivotTable1(pvtTable As PivotTable, arrDefaultHeaders As Variant)
    ' For example, we may want a time series of average sales by store types
    ' Value field
    With pvtTable.PivotFields(arrDefaultHeaders(4))
            .Orientation = xlDataField
            .Function = xlAverage
            .NumberFormat = "###.0"
    End With
    
    With pvtTable
        .PivotFields(arrDefaultHeaders(0)).Orientation = xlRowField
        .PivotFields(arrDefaultHeaders(0)).Subtotals(1) = False ' Turn off subtotals
        .PivotFields(arrDefaultHeaders(1)).Orientation = xlRowField
        .PivotFields(arrDefaultHeaders(2)).Orientation = xlPageField
        .PivotFields(arrDefaultHeaders(3)).Orientation = xlColumnField
        .ColumnGrand = False
        .RowGrand = False
        '.RowAxisLayout xlTabularRow
        '.RepeatAllLabels xlRepeatLabels
    End With
End Sub
Sub DefinePivotChart1(pvtTable As PivotTable, wsChart As Worksheet)
    Dim chtPivot As Chart, chtShape As Variant, rngTemp As Range, m As Long, n As Long
    Set rngTemp = pvtTable.DataBodyRange.Offset(1, 2)
    With rngTemp
        .Cells(1, .Columns.Count).Select
    End With
    
    For Each chtShape In wsChart.Shapes
        chtShape.Delete
    Next
    
    Set chtPivot = wsChart.Shapes.AddChart2(Style:=227, XlChartType:=xlLine).Chart
    With chtPivot
        .SetSourceData Source:=pvtTable.DataBodyRange
        .Parent.name = "TimeSeriesChart"
        .ShowAllFieldButtons = False
        .HasLegend = True
        .ChartTitle.Caption = "Time Series Plot (% over time)"
    End With
End Sub
