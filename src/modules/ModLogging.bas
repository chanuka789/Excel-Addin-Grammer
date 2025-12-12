Attribute VB_Name = "ModLogging"
'==============================================================================
' Module: ModLogging
' Description: Logging and change tracking functionality
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const LOG_WORKSHEET_NAME As String = "ChangeLog"
Private Const MAX_LOG_ENTRIES As Long = 10000

'------------------------------------------------------------------------------
' Error Record Type
'------------------------------------------------------------------------------

Public Type ErrorRecord
    CellAddress As String
    SheetName As String
    WorkbookName As String
    errorType As ModUtility.ErrorType
    OriginalText As String
    CorrectedText As String
    Severity As ModUtility.ErrorSeverity
    Category As String
    Timestamp As String
    Applied As Boolean
End Type

'------------------------------------------------------------------------------
' Change Stack for Undo Functionality
'------------------------------------------------------------------------------

Private changeStack As Collection

'------------------------------------------------------------------------------
' Initialize Logging System
'------------------------------------------------------------------------------

Public Sub InitializeLogging()
    ' Initialize change stack
    Set changeStack = New Collection

    ' Ensure ChangeLog worksheet exists
    Call EnsureLogWorksheet
End Sub

'------------------------------------------------------------------------------
' Log Event (General Purpose)
'------------------------------------------------------------------------------

Public Sub LogEvent(ByVal message As String, ByVal level As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = GetLogWorksheet()
    If ws Is Nothing Then Exit Sub

    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Prevent log from growing too large
    If lastRow > MAX_LOG_ENTRIES Then
        Call TrimLog(ws)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ' Write log entry
    ws.Cells(lastRow, 1).Value = ModUtility.GetTimestamp()
    ws.Cells(lastRow, 2).Value = level
    ws.Cells(lastRow, 3).Value = message

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Log Correction (Specific to Error Corrections)
'------------------------------------------------------------------------------

Public Sub LogCorrection(ByRef errRecord As ErrorRecord)
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = GetLogWorksheet()
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Prevent log from growing too large
    If lastRow > MAX_LOG_ENTRIES Then
        Call TrimLog(ws)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ' Write correction log
    ws.Cells(lastRow, 1).Value = errRecord.Timestamp
    ws.Cells(lastRow, 2).Value = "CORRECTION"
    ws.Cells(lastRow, 3).Value = errRecord.WorkbookName & "!" & errRecord.SheetName & "!" & errRecord.CellAddress
    ws.Cells(lastRow, 4).Value = ModUtility.ErrorTypeToString(errRecord.errorType)
    ws.Cells(lastRow, 5).Value = errRecord.OriginalText
    ws.Cells(lastRow, 6).Value = errRecord.CorrectedText
    ws.Cells(lastRow, 7).Value = ModUtility.SeverityToString(errRecord.Severity)

    ' Add to undo stack
    Call PushToUndoStack(errRecord)

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Undo Last Correction
'------------------------------------------------------------------------------

Public Function UndoLastCorrection() As Boolean
    On Error GoTo ErrorHandler

    UndoLastCorrection = False

    If changeStack.Count = 0 Then
        MsgBox "No changes to undo.", vbInformation, "Undo"
        Exit Function
    End If

    ' Get last change
    Dim lastChange As ErrorRecord
    lastChange = changeStack(changeStack.Count)

    ' Find the workbook and cell
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range

    ' Try to find the workbook
    On Error Resume Next
    Set wb = Workbooks(lastChange.WorkbookName)
    On Error GoTo ErrorHandler

    If wb Is Nothing Then
        MsgBox "Cannot undo: Workbook '" & lastChange.WorkbookName & "' is not open.", vbExclamation
        Exit Function
    End If

    Set ws = wb.Worksheets(lastChange.SheetName)
    Set cell = ws.Range(lastChange.CellAddress)

    ' Restore original text
    cell.Value = lastChange.OriginalText

    ' Remove from stack
    changeStack.Remove changeStack.Count

    ' Log the undo
    Call LogEvent("Undone correction in " & lastChange.CellAddress, "UNDO")

    UndoLastCorrection = True
    MsgBox "Last change has been undone.", vbInformation, "Undo Successful"

    Exit Function

ErrorHandler:
    MsgBox "Error during undo: " & Err.Description, vbCritical
    UndoLastCorrection = False
End Function

'------------------------------------------------------------------------------
' Undo Stack Management
'------------------------------------------------------------------------------

Private Sub PushToUndoStack(ByRef errRecord As ErrorRecord)
    ' Add to stack
    changeStack.Add errRecord

    ' Limit stack size to prevent memory issues
    If changeStack.Count > 100 Then
        changeStack.Remove 1 ' Remove oldest
    End If
End Sub

Public Function GetUndoStackSize() As Long
    GetUndoStackSize = changeStack.Count
End Function

Public Sub ClearUndoStack()
    Set changeStack = New Collection
End Sub

'------------------------------------------------------------------------------
' Helper Functions
'------------------------------------------------------------------------------

Private Function GetLogWorksheet() As Worksheet
    Set GetLogWorksheet = ModUtility.GetWorksheetByName(LOG_WORKSHEET_NAME, ThisWorkbook)
End Function

Private Sub EnsureLogWorksheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_WORKSHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new log worksheet
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = LOG_WORKSHEET_NAME

        ' Add headers
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Location"
        ws.Cells(1, 4).Value = "ErrorType"
        ws.Cells(1, 5).Value = "Original"
        ws.Cells(1, 6).Value = "Corrected"
        ws.Cells(1, 7).Value = "Severity"

        ' Format headers
        ws.Range("A1:G1").Font.Bold = True
        ws.Range("A1:G1").Interior.Color = RGB(200, 200, 200)

        ' Hide the worksheet
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

Private Sub TrimLog(ByRef ws As Worksheet)
    ' Keep only the last 5000 entries
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 5000 Then
        ' Delete old rows (keep header + last 5000)
        ws.Rows("2:" & CStr(lastRow - 5000)).Delete
    End If
End Sub

'------------------------------------------------------------------------------
' Export Log to File
'------------------------------------------------------------------------------

Public Sub ExportLogToFile(Optional ByVal filePath As String = "")
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetLogWorksheet()

    If ws Is Nothing Then
        MsgBox "No log data found.", vbInformation
        Exit Sub
    End If

    ' If no path specified, ask user
    If filePath = "" Then
        filePath = Application.GetSaveAsFilename( _
            FileFilter:="CSV Files (*.csv), *.csv", _
            Title:="Export Change Log")

        If filePath = "False" Then Exit Sub ' User cancelled
    End If

    ' Copy worksheet to new workbook and save as CSV
    Dim tempWb As Workbook
    ws.Copy
    Set tempWb = ActiveWorkbook
    tempWb.SaveAs Filename:=filePath, FileFormat:=xlCSV
    tempWb.Close SaveChanges:=False

    MsgBox "Log exported successfully to:" & vbCrLf & filePath, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error exporting log: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Clear Log
'------------------------------------------------------------------------------

Public Sub ClearLog()
    Dim ws As Worksheet
    Dim result As VbMsgBoxResult

    result = MsgBox("Are you sure you want to clear the entire change log? This cannot be undone.", _
                    vbYesNo + vbQuestion, "Clear Log")

    If result = vbNo Then Exit Sub

    Set ws = GetLogWorksheet()
    If Not ws Is Nothing Then
        ' Clear all rows except header
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        If lastRow > 1 Then
            ws.Rows("2:" & CStr(lastRow)).Delete
        End If

        MsgBox "Log cleared successfully.", vbInformation
    End If
End Sub
