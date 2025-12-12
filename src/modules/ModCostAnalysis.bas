Attribute VB_Name = "ModCostAnalysis"
'==============================================================================
' Module: ModCostAnalysis
' Description: Cost and rate validation with anomaly detection
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Analyze Costs in Range
'------------------------------------------------------------------------------

Public Sub AnalyzeCosts(ByRef targetRange As Range)
    Call CheckZeroOrNegativeRates(targetRange)
    Call DetectRateAnomalies(targetRange)
End Sub

'------------------------------------------------------------------------------
' Check for Zero or Negative Rates
'------------------------------------------------------------------------------

Private Sub CheckZeroOrNegativeRates(ByRef targetRange As Range)
    Dim cell As Range
    Dim rate As Double
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        If IsNumeric(cell.Value) Then
            rate = CDbl(cell.Value)

            ' Check for zero or negative rates
            If rate <= 0 And rate <> 0 Then ' Negative
                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etCostAnomaly
                errRecord.OriginalText = CStr(rate)
                errRecord.CorrectedText = "(Review - Negative Rate)"
                errRecord.Severity = ModUtility.esCritical
                errRecord.Category = "Negative Rate"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            ElseIf rate = 0 Then
                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etCostAnomaly
                errRecord.OriginalText = "0"
                errRecord.CorrectedText = "(Review - Zero Rate)"
                errRecord.Severity = ModUtility.esWarning
                errRecord.Category = "Zero Rate"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Detect Rate Anomalies (Outliers)
'------------------------------------------------------------------------------

Private Sub DetectRateAnomalies(ByRef targetRange As Range)
    ' Simple statistical outlier detection

    ' Collect all numeric values (potential rates)
    Dim rates() As Double
    Dim rateCount As Long
    Dim cell As Range

    ReDim rates(0 To targetRange.Cells.Count - 1)
    rateCount = 0

    For Each cell In targetRange.Cells
        If IsNumeric(cell.Value) And CDbl(cell.Value) > 0 Then
            rates(rateCount) = CDbl(cell.Value)
            rateCount = rateCount + 1
        End If
    Next cell

    If rateCount < 5 Then Exit Sub ' Need enough data points

    ' Calculate mean and standard deviation
    Dim mean As Double
    Dim stdDev As Double
    Dim i As Long
    Dim sum As Double

    sum = 0
    For i = 0 To rateCount - 1
        sum = sum + rates(i)
    Next i
    mean = sum / rateCount

    Dim variance As Double
    variance = 0
    For i = 0 To rateCount - 1
        variance = variance + (rates(i) - mean) ^ 2
    Next i
    variance = variance / rateCount
    stdDev = Sqr(variance)

    ' Flag outliers (values more than 2 std deviations from mean)
    Dim threshold As Double
    threshold = ModConfig.CostAnomalyThresholdPercent / 100 * mean

    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        If IsNumeric(cell.Value) And CDbl(cell.Value) > 0 Then
            Dim rate As Double
            rate = CDbl(cell.Value)

            ' Check if significantly higher or lower than mean
            If Abs(rate - mean) > threshold Then
                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etCostAnomaly
                errRecord.OriginalText = CStr(rate)
                errRecord.CorrectedText = "(Review - Unusual Rate, Mean: " & Format(mean, "0.00") & ")"
                errRecord.Severity = ModUtility.esInfo
                errRecord.Category = "Rate Outlier"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            End If
        End If
    Next cell
End Sub
