Attribute VB_Name = "ModFIDIC"
'==============================================================================
' Module: ModFIDIC
' Description: FIDIC clause reference validation
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const FIDIC_REFS_WORKSHEET_NAME As String = "QS_FIDICReferences"

Private fidicRefs As Collection
Private fidicLoaded As Boolean

'------------------------------------------------------------------------------
' Initialize FIDIC References
'------------------------------------------------------------------------------

Public Sub InitializeFIDIC()
    On Error GoTo ErrorHandler

    Set fidicRefs = New Collection
    fidicLoaded = False

    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(FIDIC_REFS_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then Exit Sub

    ' Load clause numbers
    Dim lastRow As Long
    Dim i As Long
    Dim clauseNum As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        clauseNum = Trim(ws.Cells(i, 1).Value)
        If Len(clauseNum) > 0 Then
            On Error Resume Next
            fidicRefs.Add clauseNum, clauseNum
            On Error GoTo ErrorHandler
        End If
    Next i

    fidicLoaded = True
    Call ModLogging.LogEvent("FIDIC references loaded: " & CStr(fidicRefs.Count), "INFO")

    Exit Sub

ErrorHandler:
    If Err.Number <> 457 Then
        Call ModLogging.LogEvent("Error loading FIDIC refs: " & Err.Description, "ERROR")
    End If
    Resume Next
End Sub

'------------------------------------------------------------------------------
' Validate FIDIC References in Range
'------------------------------------------------------------------------------

Public Sub ValidateFIDICReferences(ByRef targetRange As Range)
    If Not fidicLoaded Then Call InitializeFIDIC
    If fidicRefs.Count = 0 Then Exit Sub

    Dim cell As Range
    Dim cellText As String
    Dim clausePattern As String
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        cellText = Trim(cell.Value)

        ' Look for clause references (e.g., "Clause 1.1", "1.1.1", etc.)
        If InStr(1, cellText, "Clause", vbTextCompare) > 0 Or _
           InStr(1, cellText, "clause", vbTextCompare) > 0 Then

            ' Extract clause number (simple pattern matching)
            Dim clauseNum As String
            clauseNum = ExtractClauseNumber(cellText)

            If Len(clauseNum) > 0 Then
                If Not IsValidFIDICClause(clauseNum) Then
                    errRecord.CellAddress = cell.Address
                    errRecord.SheetName = cell.Worksheet.Name
                    errRecord.WorkbookName = cell.Worksheet.Parent.Name
                    errRecord.errorType = ModUtility.etFIDICError
                    errRecord.OriginalText = clauseNum
                    errRecord.CorrectedText = "(Verify FIDIC Clause)"
                    errRecord.Severity = ModUtility.esWarning
                    errRecord.Category = "FIDIC Reference"
                    errRecord.Timestamp = ModUtility.GetTimestamp()
                    errRecord.Applied = False

                    ModMain.g_ErrorCollection.Add errRecord
                End If
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Check if FIDIC Clause is Valid
'------------------------------------------------------------------------------

Private Function IsValidFIDICClause(ByVal clauseNum As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = fidicRefs(clauseNum)
    IsValidFIDICClause = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Extract Clause Number from Text
'------------------------------------------------------------------------------

Private Function ExtractClauseNumber(ByVal text As String) As String
    ' Simple extraction of patterns like "1.1", "1.1.1", etc.

    Dim i As Long
    Dim char As String
    Dim clauseNum As String
    Dim inNumber As Boolean

    inNumber = False
    clauseNum = ""

    For i = 1 To Len(text)
        char = Mid(text, i, 1)

        If char Like "[0-9.]" Then
            clauseNum = clauseNum & char
            inNumber = True
        ElseIf inNumber Then
            ' End of clause number
            Exit For
        End If
    Next i

    ' Clean up trailing period
    If Right(clauseNum, 1) = "." Then
        clauseNum = Left(clauseNum, Len(clauseNum) - 1)
    End If

    ExtractClauseNumber = clauseNum
End Function
