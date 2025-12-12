Attribute VB_Name = "ModQSDictionary"
'==============================================================================
' Module: ModQSDictionary
' Description: QS/Construction terminology dictionary
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const QS_DICTIONARY_WORKSHEET_NAME As String = "QS_Dictionary"

Private qsDictCache As Collection
Private qsDictLoaded As Boolean

'------------------------------------------------------------------------------
' Initialize QS Dictionary
'------------------------------------------------------------------------------

Public Sub InitializeQSDictionary()
    On Error GoTo ErrorHandler

    Set qsDictCache = New Collection
    qsDictLoaded = False

    ' Load from worksheet
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(QS_DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        Call ModLogging.LogEvent("QS Dictionary worksheet not found", "WARNING")
        Exit Sub
    End If

    ' Load terms into cache
    Dim lastRow As Long
    Dim i As Long
    Dim term As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow ' Skip header
        term = UCase(Trim(ws.Cells(i, 1).Value))
        If Len(term) > 0 Then
            On Error Resume Next
            qsDictCache.Add term, term
            On Error GoTo ErrorHandler
        End If
    Next i

    qsDictLoaded = True
    Call ModLogging.LogEvent("QS Dictionary loaded: " & CStr(qsDictCache.Count) & " terms", "INFO")

    Exit Sub

ErrorHandler:
    If Err.Number <> 457 Then ' Ignore duplicate key
        Call ModLogging.LogEvent("Error loading QS dictionary: " & Err.Description, "ERROR")
    End If
    Resume Next
End Sub

'------------------------------------------------------------------------------
' Check if Term is in QS Dictionary
'------------------------------------------------------------------------------

Public Function IsQSTerm(ByVal term As String) As Boolean
    If Not qsDictLoaded Then Call InitializeQSDictionary

    term = UCase(Trim(term))

    On Error Resume Next
    Dim temp As Variant
    temp = qsDictCache(term)
    IsQSTerm = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Get QS Term Correction
'------------------------------------------------------------------------------

Public Function GetQSTermCorrection(ByVal misspelledTerm As String) As String
    ' Find closest match in QS dictionary

    If Not qsDictLoaded Then Call InitializeQSDictionary

    ' Load worksheet for full search
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(QS_DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        GetQSTermCorrection = ""
        Exit Function
    End If

    Dim lastRow As Long
    Dim i As Long
    Dim term As String
    Dim distance As Integer
    Dim minDistance As Integer
    Dim bestMatch As String

    minDistance = 999
    bestMatch = ""

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        term = Trim(ws.Cells(i, 1).Value)

        distance = ModUtility.LevenshteinDistance(UCase(misspelledTerm), UCase(term))

        If distance < minDistance And distance <= 2 Then
            minDistance = distance
            bestMatch = term
        End If
    Next i

    GetQSTermCorrection = bestMatch
End Function

'------------------------------------------------------------------------------
' Get Standard Unit for Term
'------------------------------------------------------------------------------

Public Function GetStandardUnitForTerm(ByVal term As String) As String
    ' Returns the standard unit associated with a QS term

    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(QS_DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        GetStandardUnitForTerm = ""
        Exit Function
    End If

    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(term)) Then
            GetStandardUnitForTerm = Trim(ws.Cells(i, 6).Value) ' Column F = StandardUnit
            Exit Function
        End If
    Next i

    GetStandardUnitForTerm = ""
End Function
