Attribute VB_Name = "ModBOQAnalysis"
'==============================================================================
' Module: ModBOQAnalysis
' Description: BOQ structure and completeness analysis
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private boqInitialized As Boolean

'------------------------------------------------------------------------------
' Initialize BOQ Analysis
'------------------------------------------------------------------------------

Public Sub InitializeBOQAnalysis()
    boqInitialized = True
End Sub

'------------------------------------------------------------------------------
' Analyze BOQ Structure
'------------------------------------------------------------------------------

Public Sub AnalyzeBOQStructure(ByRef targetRange As Range)
    ' Comprehensive BOQ structure validation

    Call CheckMissingData(targetRange)
    Call ValidateCalculations(targetRange)
    Call CheckDuplicateDescriptions(targetRange)
End Sub

'------------------------------------------------------------------------------
' Check for Missing Data
'------------------------------------------------------------------------------

Public Sub CheckMissingData(ByRef targetRange As Range)
    ' Check for missing quantities, rates, units, descriptions

    Dim cell As Range
    Dim rowRange As Range
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        ' Skip header row (assuming first row)
        If cell.Row > targetRange.Row Then
            ' Check if this cell appears to be a description column (has text)
            If Len(Trim(cell.Value)) > 0 And Not IsNumeric(cell.Value) Then
                Set rowRange = Range(cell, cell.Offset(0, 4))

                ' Check for missing unit (typically next column or two columns over)
                Dim unitCell As Range
                Dim qtyCell As Range
                Dim rateCell As Range

                Set unitCell = cell.Offset(0, 1)
                Set qtyCell = cell.Offset(0, 2)
                Set rateCell = cell.Offset(0, 3)

                ' Missing Unit
                If Len(Trim(unitCell.Value)) = 0 Then
                    errRecord.CellAddress = unitCell.Address
                    errRecord.SheetName = unitCell.Worksheet.Name
                    errRecord.WorkbookName = unitCell.Worksheet.Parent.Name
                    errRecord.errorType = ModUtility.etMissingData
                    errRecord.OriginalText = "(empty)"
                    errRecord.CorrectedText = "[Unit Required]"
                    errRecord.Severity = ModUtility.esCritical
                    errRecord.Category = "Missing Unit"
                    errRecord.Timestamp = ModUtility.GetTimestamp()
                    errRecord.Applied = False

                    ModMain.g_ErrorCollection.Add errRecord
                End If

                ' Missing Quantity
                If Len(Trim(qtyCell.Value)) = 0 Or Not IsNumeric(qtyCell.Value) Then
                    errRecord.CellAddress = qtyCell.Address
                    errRecord.SheetName = qtyCell.Worksheet.Name
                    errRecord.WorkbookName = qtyCell.Worksheet.Parent.Name
                    errRecord.errorType = ModUtility.etMissingData
                    errRecord.OriginalText = "(empty)"
                    errRecord.CorrectedText = "[Quantity Required]"
                    errRecord.Severity = ModUtility.esCritical
                    errRecord.Category = "Missing Quantity"
                    errRecord.Timestamp = ModUtility.GetTimestamp()
                    errRecord.Applied = False

                    ModMain.g_ErrorCollection.Add errRecord
                End If

                ' Missing Rate
                If Len(Trim(rateCell.Value)) = 0 Or Not IsNumeric(rateCell.Value) Then
                    errRecord.CellAddress = rateCell.Address
                    errRecord.SheetName = rateCell.Worksheet.Name
                    errRecord.WorkbookName = rateCell.Worksheet.Parent.Name
                    errRecord.errorType = ModUtility.etMissingData
                    errRecord.OriginalText = "(empty)"
                    errRecord.CorrectedText = "[Rate Required]"
                    errRecord.Severity = ModUtility.esWarning
                    errRecord.Category = "Missing Rate"
                    errRecord.Timestamp = ModUtility.GetTimestamp()
                    errRecord.Applied = False

                    ModMain.g_ErrorCollection.Add errRecord
                End If
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Validate Calculations (Qty × Rate = Amount)
'------------------------------------------------------------------------------

Public Sub ValidateCalculations(ByRef targetRange As Range)
    ' Verify that Amount = Quantity × Rate

    Dim cell As Range
    Dim qtyCell As Range
    Dim rateCell As Range
    Dim amountCell As Range
    Dim qty As Double
    Dim rate As Double
    Dim amount As Double
    Dim calculated As Double
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        ' Look for numeric values that might be quantities
        If IsNumeric(cell.Value) And cell.Value <> 0 Then
            Set qtyCell = cell
            Set rateCell = cell.Offset(0, 1)
            Set amountCell = cell.Offset(0, 2)

            ' Check if all three are numeric
            If IsNumeric(qtyCell.Value) And IsNumeric(rateCell.Value) And IsNumeric(amountCell.Value) Then
                qty = CDbl(qtyCell.Value)
                rate = CDbl(rateCell.Value)
                amount = CDbl(amountCell.Value)
                calculated = qty * rate

                ' Check if calculation is correct (with small tolerance for rounding)
                If Abs(amount - calculated) > 0.01 Then
                    errRecord.CellAddress = amountCell.Address
                    errRecord.SheetName = amountCell.Worksheet.Name
                    errRecord.WorkbookName = amountCell.Worksheet.Parent.Name
                    errRecord.errorType = ModUtility.etCalculationError
                    errRecord.OriginalText = CStr(amount)
                    errRecord.CorrectedText = Format(calculated, "0.00")
                    errRecord.Severity = ModUtility.esCritical
                    errRecord.Category = "Calculation Error"
                    errRecord.Timestamp = ModUtility.GetTimestamp()
                    errRecord.Applied = False

                    ModMain.g_ErrorCollection.Add errRecord
                End If
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Check for Duplicate Descriptions
'------------------------------------------------------------------------------

Public Sub CheckDuplicateDescriptions(ByRef targetRange As Range)
    ' Find duplicate or very similar BOQ descriptions

    Dim descriptions As Collection
    Set descriptions = New Collection

    Dim cell As Range
    Dim desc As String
    Dim existingDesc As Variant
    Dim similarity As Integer
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        desc = Trim(cell.Value)

        ' Only check text cells with substantial content
        If Len(desc) > 10 And Not IsNumeric(desc) Then
            ' Check for exact duplicates
            On Error Resume Next
            Dim testVal As Variant
            testVal = descriptions(UCase(desc))
            If Err.Number = 0 Then
                ' Exact duplicate found
                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etDescriptionError
                errRecord.OriginalText = desc
                errRecord.CorrectedText = "(Duplicate - Review)"
                errRecord.Severity = ModUtility.esWarning
                errRecord.Category = "Duplicate Description"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            Else
                ' Add to collection
                descriptions.Add cell.Address, UCase(desc)
            End If
            On Error GoTo 0
        End If
    Next cell
End Sub
