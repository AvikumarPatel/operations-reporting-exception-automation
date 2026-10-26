Attribute VB_Name = "OperationsAutomation"
Option Explicit

Public Sub ProcessOperationsReport()

    Dim wsRaw As Worksheet
    Dim wsKPI As Worksheet
    Dim wsExceptions As Worksheet
    Dim wsConfig As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim exceptionRow As Long

    Dim orderID As String
    Dim warehouse As String
    Dim employeeID As String

    Dim ordersProcessed As Double
    Dim unitsPicked As Double
    Dim errors As Double
    Dim processingMinutes As Double
    Dim lateOrders As Double
    Dim returnsCount As Double

    Dim errorRate As Double
    Dim lateRate As Double
    Dim returnRate As Double
    Dim ordersPerHour As Double

    Dim errorThreshold As Double
    Dim lateThreshold As Double
    Dim productivityThreshold As Double
    Dim returnThreshold As Double

    Dim exceptionFlags As String
    Dim dataQualityFlags As String

    Dim orderDictionary As Object

    Dim totalOrders As Double
    Dim totalUnits As Double
    Dim totalErrors As Double
    Dim totalLate As Double
    Dim totalReturns As Double
    Dim totalProcessingMinutes As Double

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsRaw = ThisWorkbook.Worksheets("Raw_Data")
    Set wsKPI = ThisWorkbook.Worksheets("KPI_Summary")
    Set wsExceptions = ThisWorkbook.Worksheets("Exceptions")
    Set wsConfig = ThisWorkbook.Worksheets("Config")

    Set orderDictionary = CreateObject("Scripting.Dictionary")

    ' Read configurable thresholds from Config sheet.
    errorThreshold = wsConfig.Range("B2").Value
    lateThreshold = wsConfig.Range("B3").Value
    productivityThreshold = wsConfig.Range("B4").Value
    returnThreshold = wsConfig.Range("B5").Value

    If WorksheetFunction.CountA(wsRaw.Cells) = 0 Then
        GoTo CleanExit
    End If

    lastRow = wsRaw.Cells.Find( _
        What:="*", _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).Row

    ' Calculated output columns.
    wsRaw.Cells(1, 12).Value = "ErrorRate"
    wsRaw.Cells(1, 13).Value = "LateOrderRate"
    wsRaw.Cells(1, 14).Value = "ReturnRate"
    wsRaw.Cells(1, 15).Value = "OrdersPerHour"
    wsRaw.Cells(1, 16).Value = "ExceptionFlags"
    wsRaw.Cells(1, 17).Value = "DataQualityIssue"

    wsRaw.Range("L1:Q1").Font.Bold = True
    wsRaw.Range("L1:Q1").Interior.Color = RGB(31, 78, 121)
    wsRaw.Range("L1:Q1").Font.Color = RGB(255, 255, 255)

    If lastRow >= 2 Then
        wsRaw.Range("L2:Q" & lastRow).ClearContents
    End If

    ' Rebuild Exceptions sheet for the current run.
    wsExceptions.Cells.Clear
    wsRaw.Range("A1:Q1").Copy Destination:=wsExceptions.Range("A1")
    exceptionRow = 2

    For i = 2 To lastRow

        exceptionFlags = ""
        dataQualityFlags = ""

        orderID = Trim(CStr(wsRaw.Cells(i, 1).Value))
        warehouse = Trim(CStr(wsRaw.Cells(i, 3).Value))
        employeeID = Trim(CStr(wsRaw.Cells(i, 4).Value))

        wsRaw.Cells(i, 1).Value = orderID
        wsRaw.Cells(i, 3).Value = warehouse
        wsRaw.Cells(i, 4).Value = employeeID

        ' Data-quality checks.
        If orderID = "" Then
            dataQualityFlags = AppendFlag(dataQualityFlags, "MISSING_ORDER_ID")
        ElseIf orderDictionary.Exists(orderID) Then
            dataQualityFlags = AppendFlag(dataQualityFlags, "DUPLICATE_ORDER_ID")
        Else
            orderDictionary.Add orderID, True
        End If

        If warehouse = "" Then
            dataQualityFlags = AppendFlag(dataQualityFlags, "MISSING_WAREHOUSE")
        End If

        If employeeID = "" Then
            dataQualityFlags = AppendFlag(dataQualityFlags, "MISSING_EMPLOYEE_ID")
        End If

        ' Numeric validation and extraction.
        If IsNumeric(wsRaw.Cells(i, 6).Value) Then
            ordersProcessed = CDbl(wsRaw.Cells(i, 6).Value)
        Else
            ordersProcessed = 0
            dataQualityFlags = AppendFlag(dataQualityFlags, "INVALID_ORDERS_PROCESSED")
        End If

        If IsNumeric(wsRaw.Cells(i, 7).Value) Then
            unitsPicked = CDbl(wsRaw.Cells(i, 7).Value)
        Else
            unitsPicked = 0
        End If

        If IsNumeric(wsRaw.Cells(i, 8).Value) Then
            errors = CDbl(wsRaw.Cells(i, 8).Value)
        Else
            errors = 0
        End If

        If IsNumeric(wsRaw.Cells(i, 9).Value) Then
            processingMinutes = CDbl(wsRaw.Cells(i, 9).Value)
        Else
            processingMinutes = 0
            dataQualityFlags = AppendFlag(dataQualityFlags, "INVALID_PROCESSING_TIME")
        End If

        If IsNumeric(wsRaw.Cells(i, 10).Value) Then
            lateOrders = CDbl(wsRaw.Cells(i, 10).Value)
        Else
            lateOrders = 0
        End If

        If IsNumeric(wsRaw.Cells(i, 11).Value) Then
            returnsCount = CDbl(wsRaw.Cells(i, 11).Value)
        Else
            returnsCount = 0
        End If

        ' KPI calculations.
        If ordersProcessed > 0 Then
            errorRate = errors / ordersProcessed
            lateRate = lateOrders / ordersProcessed
            returnRate = returnsCount / ordersProcessed
        Else
            errorRate = 0
            lateRate = 0
            returnRate = 0
            dataQualityFlags = AppendFlag(dataQualityFlags, "ZERO_ORDERS_PROCESSED")
        End If

        If processingMinutes > 0 Then
            ordersPerHour = ordersProcessed / (processingMinutes / 60)
        Else
            ordersPerHour = 0
            dataQualityFlags = AppendFlag(dataQualityFlags, "ZERO_PROCESSING_TIME")
        End If

        ' Configurable operational exception rules.
        If errorRate > errorThreshold Then
            exceptionFlags = AppendFlag(exceptionFlags, "HIGH_ERROR_RATE")
        End If

        If lateRate > lateThreshold Then
            exceptionFlags = AppendFlag(exceptionFlags, "HIGH_LATE_ORDER_RATE")
        End If

        If ordersPerHour < productivityThreshold And ordersProcessed > 0 Then
            exceptionFlags = AppendFlag(exceptionFlags, "LOW_PRODUCTIVITY")
        End If

        If returnRate > returnThreshold Then
            exceptionFlags = AppendFlag(exceptionFlags, "HIGH_RETURN_RATE")
        End If

        wsRaw.Cells(i, 12).Value = errorRate
        wsRaw.Cells(i, 13).Value = lateRate
        wsRaw.Cells(i, 14).Value = returnRate
        wsRaw.Cells(i, 15).Value = ordersPerHour
        wsRaw.Cells(i, 16).Value = exceptionFlags
        wsRaw.Cells(i, 17).Value = dataQualityFlags

        If exceptionFlags <> "" Or dataQualityFlags <> "" Then
            wsRaw.Range("A" & i & ":Q" & i).Copy _
                Destination:=wsExceptions.Range("A" & exceptionRow)
            exceptionRow = exceptionRow + 1
        End If

        totalOrders = totalOrders + ordersProcessed
        totalUnits = totalUnits + unitsPicked
        totalErrors = totalErrors + errors
        totalLate = totalLate + lateOrders
        totalReturns = totalReturns + returnsCount
        totalProcessingMinutes = totalProcessingMinutes + processingMinutes

    Next i

    wsRaw.Range("L2:N" & lastRow).NumberFormat = "0.00%"
    wsRaw.Range("O2:O" & lastRow).NumberFormat = "0.00"
    wsRaw.Columns("A:Q").AutoFit
    wsExceptions.Columns("A:Q").AutoFit

    ' KPI summary output.
    wsKPI.Cells.Clear
    wsKPI.Range("A1").Value = "Operations Performance Summary"
    wsKPI.Range("A1:B1").Merge
    wsKPI.Range("A1:B1").Font.Bold = True
    wsKPI.Range("A1:B1").Font.Size = 16

    wsKPI.Range("A3").Value = "Metric"
    wsKPI.Range("B3").Value = "Value"
    wsKPI.Range("A3:B3").Font.Bold = True

    wsKPI.Range("A4").Value = "Total Records"
    wsKPI.Range("B4").Value = lastRow - 1

    wsKPI.Range("A5").Value = "Unique Order IDs"
    wsKPI.Range("B5").Value = orderDictionary.Count

    wsKPI.Range("A6").Value = "Total Orders Processed"
    wsKPI.Range("B6").Value = totalOrders

    wsKPI.Range("A7").Value = "Total Units Picked"
    wsKPI.Range("B7").Value = totalUnits

    wsKPI.Range("A8").Value = "Overall Error Rate"
    If totalOrders > 0 Then wsKPI.Range("B8").Value = totalErrors / totalOrders

    wsKPI.Range("A9").Value = "Late Order Rate"
    If totalOrders > 0 Then wsKPI.Range("B9").Value = totalLate / totalOrders

    wsKPI.Range("A10").Value = "Return Rate"
    If totalOrders > 0 Then wsKPI.Range("B10").Value = totalReturns / totalOrders

    wsKPI.Range("A11").Value = "Overall Orders / Hour"
    If totalProcessingMinutes > 0 Then
        wsKPI.Range("B11").Value = totalOrders / (totalProcessingMinutes / 60)
    End If

    wsKPI.Range("A12").Value = "Exception Records"
    wsKPI.Range("B12").Value = exceptionRow - 2

    wsKPI.Range("B8:B10").NumberFormat = "0.00%"
    wsKPI.Range("B11").NumberFormat = "0.00"
    wsKPI.Columns("A:B").AutoFit

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Function AppendFlag(ByVal existingFlags As String, ByVal newFlag As String) As String
    If existingFlags = "" Then
        AppendFlag = newFlag
    Else
        AppendFlag = existingFlags & "; " & newFlag
    End If
End Function
