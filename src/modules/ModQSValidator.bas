Attribute VB_Name = "ModQSValidator"
'==============================================================================
' Module: ModQSValidator
' Description: Main QS validation orchestrator
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private qsInitialized As Boolean

'------------------------------------------------------------------------------
' Initialize QS Module
'------------------------------------------------------------------------------

Public Sub InitializeQS()
    On Error Resume Next

    ' Initialize sub-modules
    Call ModQSDictionary.InitializeQSDictionary
    Call ModUnitValidator.InitializeUnitValidator
    Call ModBOQAnalysis.InitializeBOQAnalysis

    qsInitialized = True
    Call ModLogging.LogEvent("QS modules initialized", "INFO")
End Sub

'------------------------------------------------------------------------------
' Main QS Scanning Function
'------------------------------------------------------------------------------

Public Sub ScanRangeForQSErrors(ByRef targetRange As Range)
    On Error Resume Next

    If Not qsInitialized Then Call InitializeQS

    Call ModUtility.ShowProgressMessage("Performing QS validation...")

    ' BOQ Structure Analysis
    If ModConfig.EnableBOQAnalysis Then
        Call ModBOQAnalysis.AnalyzeBOQStructure(targetRange)
    End If

    ' Unit Validation
    If ModConfig.EnableUnitValidation Then
        Call ModUnitValidator.ValidateUnits(targetRange)
    End If

    ' Cost Analysis
    If ModConfig.EnableCostAnalysis Then
        Call ModCostAnalysis.AnalyzeCosts(targetRange)
    End If

    ' Description Analysis
    Call ModDescriptionAnalysis.AnalyzeDescriptions(targetRange)

    ' FIDIC Validation (if enabled)
    If ModConfig.EnableFIDICValidation Then
        Call ModFIDIC.ValidateFIDICReferences(targetRange)
    End If

    Call ModLogging.LogEvent("QS validation completed", "INFO")
End Sub

'------------------------------------------------------------------------------
' Quick QS Check (Data Completeness Only)
'------------------------------------------------------------------------------

Public Sub QuickQSCheck(ByRef targetRange As Range)
    ' Fast check for missing data only

    Call ModUtility.ShowProgressMessage("Quick QS check...")

    ' Check for missing critical fields
    Call ModBOQAnalysis.CheckMissingData(targetRange)

    Call ModLogging.LogEvent("Quick QS check completed", "INFO")
End Sub

'------------------------------------------------------------------------------
' Validate BOQ Headers
'------------------------------------------------------------------------------

Public Function ValidateBOQHeaders(ByRef headerRange As Range) As Boolean
    ' Check if range contains required BOQ headers

    Dim requiredHeaders As Variant
    requiredHeaders = Array("Description", "Unit", "Quantity", "Rate", "Amount")

    Dim i As Long
    Dim found As Boolean
    Dim cell As Range

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        found = False
        For Each cell In headerRange.Cells
            If UCase(Trim(cell.Value)) = UCase(requiredHeaders(i)) Then
                found = True
                Exit For
            End If
        Next cell

        If Not found Then
            ValidateBOQHeaders = False
            Exit Function
        End If
    Next i

    ValidateBOQHeaders = True
End Function
