Attribute VB_Name = "ModMain"
'==============================================================================
' Module: ModMain
' Description: Main entry points and orchestration for the add-in
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Public Const ADDIN_VERSION As String = "1.0.0"
Public Const ADDIN_NAME As String = "Grammar & QS Checker"

' Global error collection
Public g_ErrorCollection As Collection

'==============================================================================
' RIBBON BUTTON CALLBACKS
'==============================================================================

'------------------------------------------------------------------------------
' Check Now Button (Main Entry Point)
'------------------------------------------------------------------------------

Public Sub CheckNow_Click(control As IRibbonControl)
    ' Main button - checks both spelling/grammar and QS if enabled
    Call CheckActiveWorkbook(checkSpelling:=True, checkGrammar:=True, checkQS:=True)
End Sub

'------------------------------------------------------------------------------
' Check Spelling Only
'------------------------------------------------------------------------------

Public Sub CheckSpelling_Click(control As IRibbonControl)
    Call CheckActiveWorkbook(checkSpelling:=True, checkGrammar:=False, checkQS:=False)
End Sub

'------------------------------------------------------------------------------
' Check Grammar Only
'------------------------------------------------------------------------------

Public Sub CheckGrammar_Click(control As IRibbonControl)
    Call CheckActiveWorkbook(checkSpelling:=False, checkGrammar:=True, checkQS:=False)
End Sub

'------------------------------------------------------------------------------
' Check QS Only
'------------------------------------------------------------------------------

Public Sub CheckQS_Click(control As IRibbonControl)
    Call CheckActiveWorkbook(checkSpelling:=False, checkGrammar:=False, checkQS:=True)
End Sub

'------------------------------------------------------------------------------
' Settings Button
'------------------------------------------------------------------------------

Public Sub ShowSettings_Click(control As IRibbonControl)
    ' Show settings form
    ' frmSettings.Show ' Uncomment when form is created
    MsgBox "Settings dialog will be shown here", vbInformation, "Settings"
End Sub

'------------------------------------------------------------------------------
' Help Button
'------------------------------------------------------------------------------

Public Sub ShowHelp_Click(control As IRibbonControl)
    Dim helpMessage As String
    helpMessage = ADDIN_NAME & " v" & ADDIN_VERSION & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Features:" & vbCrLf
    helpMessage = helpMessage & "- Spelling Check" & vbCrLf
    helpMessage = helpMessage & "- Grammar Check" & vbCrLf
    helpMessage = helpMessage & "- QS/BOQ Validation" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Select cells or ranges and click 'Check Now' to start." & vbCrLf & vbCrLf
    helpMessage = helpMessage & "For detailed help, see the User Guide."

    MsgBox helpMessage, vbInformation, "Help - " & ADDIN_NAME
End Sub

'==============================================================================
' MAIN CHECKING FUNCTION
'==============================================================================

Public Sub CheckActiveWorkbook(Optional ByVal checkSpelling As Boolean = True, _
                               Optional ByVal checkGrammar As Boolean = True, _
                               Optional ByVal checkQS As Boolean = False)
    On Error GoTo ErrorHandler

    ' Validate that a workbook is open
    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a workbook first.", vbExclamation, "No Workbook Open"
        Exit Sub
    End If

    ' Don't check the add-in itself
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Cannot check the add-in file itself. Please open a different workbook.", _
               vbExclamation, "Invalid Target"
        Exit Sub
    End If

    ' Get selection or used range
    Dim targetRange As Range
    Set targetRange = GetTargetRange()

    If targetRange Is Nothing Then
        MsgBox "No valid range to check.", vbExclamation
        Exit Sub
    End If

    ' Validate range size
    If Not ModUtility.IsValidRangeForCheck(targetRange) Then
        MsgBox "Selected range is too large. Please select a smaller range.", vbExclamation
        Exit Sub
    End If

    ' Initialize error collection
    Set g_ErrorCollection = New Collection

    ' Show progress
    Application.ScreenUpdating = False
    Call ModUtility.ShowProgressMessage("Scanning " & targetRange.Cells.Count & " cells...")

    ' Perform checks based on parameters
    If checkSpelling Or checkGrammar Then
        Call ScanRangeForSpellingAndGrammar(targetRange, checkSpelling, checkGrammar)
    End If

    If checkQS And ModConfig.EnableQSValidation Then
        Call ModQSValidator.ScanRangeForQSErrors(targetRange)
    End If

    ' Hide progress
    Call ModUtility.ClearProgressMessage()
    Application.ScreenUpdating = True

    ' Show results
    Call ShowResults

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call ModUtility.ClearProgressMessage()
    MsgBox "Error during check: " & Err.Description, vbCritical
End Sub

'==============================================================================
' SCANNING FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' Scan Range for Spelling and Grammar Errors
'------------------------------------------------------------------------------

Private Sub ScanRangeForSpellingAndGrammar(ByRef targetRange As Range, _
                                           ByVal checkSpelling As Boolean, _
                                           ByVal checkGrammar As Boolean)
    Dim cell As Range
    Dim cellText As String
    Dim cellCount As Long
    Dim processedCount As Long

    cellCount = targetRange.Cells.Count
    processedCount = 0

    For Each cell In targetRange.Cells
        processedCount = processedCount + 1

        ' Update progress every 100 cells
        If processedCount Mod 100 = 0 Then
            Call ModUtility.ShowProgressMessage("Processing: " & processedCount & " of " & cellCount)
        End If

        ' Get cell text
        cellText = ModUtility.GetCellValue(cell)

        ' Skip empty cells and very short text
        If Len(Trim(cellText)) > 1 Then
            ' Check spelling
            If checkSpelling And ModConfig.EnableSpellingCheck Then
                Call CheckCellSpelling(cell, cellText)
            End If

            ' Check grammar
            If checkGrammar And ModConfig.EnableGrammarCheck Then
                Call CheckCellGrammar(cell, cellText)
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Check Single Cell for Spelling Errors
'------------------------------------------------------------------------------

Private Sub CheckCellSpelling(ByRef cell As Range, ByVal cellText As String)
    Dim misspelledWords As Collection
    Set misspelledWords = ModSpelling.CheckSpelling(cellText)

    If misspelledWords.Count > 0 Then
        ' Create error record for each misspelled word
        Dim i As Long
        Dim wordInfo As Object
        Dim errRecord As ModLogging.ErrorRecord

        For i = 1 To misspelledWords.Count
            Set wordInfo = misspelledWords(i)

            errRecord.CellAddress = cell.Address
            errRecord.SheetName = cell.Worksheet.Name
            errRecord.WorkbookName = cell.Worksheet.Parent.Name
            errRecord.errorType = ModUtility.etSpelling
            errRecord.OriginalText = wordInfo("Word")
            errRecord.CorrectedText = "" ' Will be filled when user selects suggestion
            errRecord.Severity = ModUtility.esWarning
            errRecord.Category = "Spelling"
            errRecord.Timestamp = ModUtility.GetTimestamp()
            errRecord.Applied = False

            g_ErrorCollection.Add errRecord
        Next i
    End If
End Sub

'------------------------------------------------------------------------------
' Check Single Cell for Grammar Errors
'------------------------------------------------------------------------------

Private Sub CheckCellGrammar(ByRef cell As Range, ByVal cellText As String)
    Dim grammarErrors As Collection
    Set grammarErrors = ModGrammar.CheckGrammar(cellText)

    If grammarErrors.Count > 0 Then
        ' Create error record for each grammar error
        Dim i As Long
        Dim errorInfo As Object
        Dim errRecord As ModLogging.ErrorRecord

        For i = 1 To grammarErrors.Count
            Set errorInfo = grammarErrors(i)

            errRecord.CellAddress = cell.Address
            errRecord.SheetName = cell.Worksheet.Name
            errRecord.WorkbookName = cell.Worksheet.Parent.Name
            errRecord.errorType = ModUtility.etGrammar
            errRecord.OriginalText = errorInfo("Pattern")
            errRecord.CorrectedText = errorInfo("Replacement")
            errRecord.Severity = errorInfo("Severity")
            errRecord.Category = errorInfo("Category")
            errRecord.Timestamp = ModUtility.GetTimestamp()
            errRecord.Applied = False

            g_ErrorCollection.Add errRecord
        Next i
    End If
End Sub

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' Get Target Range (Selection or Used Range)
'------------------------------------------------------------------------------

Private Function GetTargetRange() As Range
    On Error Resume Next

    ' If user has selected cells, use selection
    If Not Selection Is Nothing Then
        If TypeName(Selection) = "Range" Then
            Set GetTargetRange = Selection
            Exit Function
        End If
    End If

    ' Otherwise use entire used range of active sheet
    Set GetTargetRange = ActiveSheet.UsedRange

    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Show Results Dialog
'------------------------------------------------------------------------------

Private Sub ShowResults()
    If g_ErrorCollection Is Nothing Then
        MsgBox "No errors found!", vbInformation, "Check Complete"
        Exit Sub
    End If

    If g_ErrorCollection.Count = 0 Then
        MsgBox "No errors found!", vbInformation, "Check Complete"
    Else
        ' Show results form
        ' frmResults.Show ' Uncomment when form is created

        ' Temporary: Show message box with count
        Dim msg As String
        msg = "Found " & g_ErrorCollection.Count & " error(s):" & vbCrLf & vbCrLf

        ' Show first few errors
        Dim i As Long
        Dim errRecord As ModLogging.ErrorRecord

        For i = 1 To Application.WorksheetFunction.Min(5, g_ErrorCollection.Count)
            errRecord = g_ErrorCollection(i)
            msg = msg & ModUtility.ErrorTypeToString(errRecord.errorType) & ": "
            msg = msg & errRecord.OriginalText
            If Len(errRecord.CorrectedText) > 0 Then
                msg = msg & " → " & errRecord.CorrectedText
            End If
            msg = msg & " (" & errRecord.CellAddress & ")" & vbCrLf
        Next i

        If g_ErrorCollection.Count > 5 Then
            msg = msg & vbCrLf & "... and " & (g_ErrorCollection.Count - 5) & " more"
        End If

        MsgBox msg, vbInformation, "Check Results"
    End If
End Sub

'==============================================================================
' INITIALIZATION
'==============================================================================

Public Sub InitializeAddIn()
    On Error Resume Next

    ' Load configuration
    Call ModConfig.LoadSettings

    ' Initialize logging
    Call ModLogging.InitializeLogging

    ' Initialize spelling dictionary
    Call ModSpelling.InitializeDictionary

    ' Initialize grammar rules
    Call ModGrammar.InitializeGrammarRules

    ' Initialize QS modules if enabled
    If ModConfig.EnableQSValidation Then
        Call ModQSValidator.InitializeQS
    End If

    Call ModLogging.LogEvent("Add-in initialized successfully", "INFO")
End Sub

Public Sub OnInstall()
    Call ModConfig.CreateDefaultSettings
    MsgBox ADDIN_NAME & " v" & ADDIN_VERSION & " installed successfully!" & vbCrLf & vbCrLf & _
           "Look for buttons in the Excel ribbon.", vbInformation, "Installation Complete"
End Sub

Public Sub OnUninstall()
    MsgBox ADDIN_NAME & " has been uninstalled.", vbInformation, "Uninstall"
End Sub
