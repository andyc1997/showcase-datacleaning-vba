Option Explicit
' Test if the given name exists in any current table
Function IsTableExist(loTables As ListObjects, sTableName As String) As Boolean
    Dim loTable As ListObject
    IsTableExist = False
    For Each loTable In loTables
        If loTable.name = sTableName Then IsTableExist = True
    Next
End Function
' Test if the data has pattern like Nov '22
Function MonthlyPattern(sDate As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[A-Z][a-z][a-z] \'\d\d"
    If regex.Test(sDate) Then MonthlyPattern = True Else MonthlyPattern = False
    Set regex = Nothing
End Function
' Cannot add UDT to VBA collection object, use class module!
' https://stackoverflow.com/questions/24954454/user-defined-types-in-arrays-collections-and-for-each-loops
Sub GetData(sSheetName As String, sRngName As String, cltSales As Collection)
    Dim rngData As Range, cntTimeSlice As Long, cntStoreType As Long, curRangeColumn As Integer, sDate As String
    Dim oMonthSales As clsMonthlySales
    
    Set rngData = ThisWorkbook.Worksheets(sSheetName).Range(sRngName)
    cntTimeSlice = rngData.Columns.Count
    cntStoreType = rngData.Rows.Count
    
    #If cDebugMode = 1 Then
        Debug.Print cntTimeSlice
        Debug.Print cntStoreType
    #End If
    
    For curRangeColumn = 1 To cntTimeSlice
        sDate = rngData(1, curRangeColumn).Offset(-1, 0) ' Go to the above cell
        If MonthlyPattern(sDate) Then
            Set oMonthSales = New clsMonthlySales
            With oMonthSales
                .Init 6 ' Six type of stores
                .MonthYear = sDate
                .UpdateArray 1, "same_store", "Net Sales", rngData(1, curRangeColumn)
                .UpdateArray 2, "same_store", "Customer Numbers", rngData(2, curRangeColumn)
                .UpdateArray 3, "same_store", "Average Purchases", rngData(3, curRangeColumn)
                .UpdateArray 4, "own_store", "Net Sales", rngData(4, curRangeColumn)
                .UpdateArray 5, "own_store", "Customer Numbers", rngData(5, curRangeColumn)
                .UpdateArray 6, "own_store", "Average Purchases", rngData(6, curRangeColumn)
            End With
            cltSales.Add oMonthSales
        Else
            #If cDebugMode = 1 Then
                Debug.Print sDate
            #End If
        End If
    Next
End Sub
Sub WriteTable(sSheetName As String, sTableName As String, cltSales As Collection, arrDefaultHeaders As Variant)
    Dim lrRecord As ListRow, loTables As ListObjects, loTable As ListObject, bExist As Boolean, rngHeaders As Range, _
        oItem As clsMonthlySales, n As Integer
    
    Set loTables = ThisWorkbook.Worksheets(sSheetName).ListObjects
    bExist = IsTableExist(loTables, sTableName) ' Check any table has the given name
    
    If bExist Then
        Set loTable = loTables(sTableName)
        loTable.DataBodyRange.Delete
    Else ' Clear all existing table(s) in the worksheet and assign column headers for a new table
        For Each loTable In loTables
            loTable.Delete
        Next
        Set rngHeaders = ThisWorkbook.Worksheets(sSheetName).Range("A1").Resize(1, UBound(arrDefaultHeaders) + 1)
        rngHeaders.value = arrDefaultHeaders
        Set loTable = loTables.Add(SourceType:=xlSrcRange, Source:=rngHeaders, xlListObjecthasheaders:=xlYes, Destination:=rngHeaders)
        loTable.name = sTableName
    End If
    
    With loTable ' Update each entry of the table
        For Each oItem In cltSales
            For n = 1 To 6
                Set lrRecord = .ListRows.Add
                lrRecord.Range(.ListColumns(arrDefaultHeaders(0)).Index) = Right(Trim(oItem.MonthYear), 2)
                lrRecord.Range(.ListColumns(arrDefaultHeaders(1)).Index) = Left(Trim(oItem.MonthYear), 3)
                lrRecord.Range(.ListColumns(arrDefaultHeaders(2)).Index) = oItem.StoreType(n)
                lrRecord.Range(.ListColumns(arrDefaultHeaders(3)).Index) = oItem.InternalName(n)
                lrRecord.Range(.ListColumns(arrDefaultHeaders(4)).Index) = oItem.InternalValue(n)
            Next
        Next
    End With
End Sub
Sub UpdatePivotTable(sSrcSheetName As String, sDstSheetName As String, sTableName As String, arrDefaultHeaders As Variant)
    Dim wsSource As Worksheet, wsDestination As Worksheet
    Dim pvtTable As PivotTable, pvtTable1 As PivotTable, pvtTables As PivotTables, pvtCache As PivotCache
    
    ' Clear all existing pivot tables
    Set wsDestination = ThisWorkbook.Worksheets(sDstSheetName)
    Set pvtTables = wsDestination.PivotTables
    For Each pvtTable In pvtTables
        pvtTable.TableRange2.Clear
    Next
    
    ' The common pivot cache for all pivot tables in this sheet
    Set wsSource = ThisWorkbook.Worksheets(sSrcSheetName)
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsSource.ListObjects(sTableName))
    
    ' Pivot Table 1 (Just a sample)
    Set pvtTable1 = pvtCache.CreatePivotTable(TableDestination:=wsDestination.Range("A1"), TableName:="TimeSeries")
    Call Interface.DefinePivotTable1(pvtTable1, arrDefaultHeaders)
    Call Interface.DefinePivotChart1(pvtTable1, wsDestination)
End Sub
