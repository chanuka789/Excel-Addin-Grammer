Attribute VB_Name = "ModUtility"
'==============================================================================
' Module: ModUtility
' Description: Utility and helper functions used across the add-in
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' String Manipulation Functions
'------------------------------------------------------------------------------

' Trim and normalize whitespace
Public Function NormalizeText(ByVal inputText As String) As String
    Dim result As String
    result = Trim(inputText)

    ' Replace multiple spaces with single space
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    NormalizeText = result
End Function

' Split text into words
Public Function SplitIntoWords(ByVal inputText As String) As Variant
    Dim normalized As String
    normalized = NormalizeText(inputText)

    ' Remove punctuation for word splitting
    Dim cleanText As String
    cleanText = RemovePunctuation(normalized)

    If Len(cleanText) > 0 Then
        SplitIntoWords = Split(cleanText, " ")
    Else
        SplitIntoWords = Array()
    End If
End Function

' Remove punctuation from text
Public Function RemovePunctuation(ByVal inputText As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String

    result = ""
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        ' Keep alphanumeric and spaces
        If char Like "[A-Za-z0-9 ]" Then
            result = result & char
        End If
    Next i

    RemovePunctuation = result
End Function

' Check if string is alphanumeric
Public Function IsAlphanumeric(ByVal inputText As String) As Boolean
    Dim i As Long
    Dim char As String

    IsAlphanumeric = True
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        If Not char Like "[A-Za-z0-9]" Then
            IsAlphanumeric = False
            Exit Function
        End If
    Next i
End Function

' Check if string is numeric
Public Function IsNumericString(ByVal inputText As String) As Boolean
    On Error Resume Next
    Dim testVal As Double
    testVal = CDbl(inputText)
    IsNumericString = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Levenshtein Distance (for spelling suggestions)
'------------------------------------------------------------------------------

' Calculate edit distance between two words
Public Function LevenshteinDistance(ByVal word1 As String, ByVal word2 As String) As Integer
    Dim len1 As Integer, len2 As Integer
    Dim i As Integer, j As Integer
    Dim cost As Integer
    Dim d() As Integer

    len1 = Len(word1)
    len2 = Len(word2)

    ' Handle edge cases
    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If
    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ' Initialize matrix
    ReDim d(0 To len1, 0 To len2)

    For i = 0 To len1
        d(i, 0) = i
    Next i

    For j = 0 To len2
        d(0, j) = j
    Next j

    ' Calculate distances
    For i = 1 To len1
        For j = 1 To len2
            If Mid(word1, i, 1) = Mid(word2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            d(i, j) = Application.WorksheetFunction.Min( _
                d(i - 1, j) + 1, _
                d(i, j - 1) + 1, _
                d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(len1, len2)
End Function

'------------------------------------------------------------------------------
' Collection Helpers
'------------------------------------------------------------------------------

' Check if item exists in collection
Public Function CollectionContains(ByRef col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionContains = (Err.Number = 0)
    On Error GoTo 0
End Function

' Safe add to collection with key
Public Sub CollectionAddSafe(ByRef col As Collection, ByVal item As Variant, ByVal key As String)
    On Error Resume Next
    col.Add item, key
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Worksheet Helpers
'------------------------------------------------------------------------------

' Get worksheet by name (safely)
Public Function GetWorksheetByName(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set GetWorksheetByName = wb.Worksheets(wsName)
    On Error GoTo 0
End Function

' Check if worksheet exists
Public Function WorksheetExists(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    WorksheetExists = Not (wb.Worksheets(wsName) Is Nothing)
    On Error GoTo 0
End Function

' Get cell value safely
Public Function GetCellValue(ByVal cell As Range) As String
    On Error Resume Next
    If Not cell.HasFormula Then
        GetCellValue = CStr(cell.Value)
    Else
        GetCellValue = "" ' Skip formula cells for text checking
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Range Helpers
'------------------------------------------------------------------------------

' Check if range is valid for checking
Public Function IsValidRangeForCheck(ByVal rng As Range) As Boolean
    IsValidRangeForCheck = False

    If rng Is Nothing Then Exit Function
    If rng.Cells.Count > 100000 Then Exit Function ' Safety limit

    IsValidRangeForCheck = True
End Function

' Get used range in worksheet
Public Function GetUsedRangeForCheck(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set GetUsedRangeForCheck = ws.UsedRange
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Date/Time Helpers
'------------------------------------------------------------------------------

' Get current timestamp
Public Function GetTimestamp() As String
    GetTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' Get date only
Public Function GetDateStamp() As String
    GetDateStamp = Format(Date, "yyyy-mm-dd")
End Function

'------------------------------------------------------------------------------
' File Path Helpers
'------------------------------------------------------------------------------

' Get add-in installation path
Public Function GetAddInPath() As String
    GetAddInPath = ThisWorkbook.Path
End Function

' Ensure path ends with separator
Public Function EnsurePathSeparator(ByVal path As String) As String
    If Right(path, 1) <> Application.PathSeparator Then
        EnsurePathSeparator = path & Application.PathSeparator
    Else
        EnsurePathSeparator = path
    End If
End Function

'------------------------------------------------------------------------------
' Error Type Enums
'------------------------------------------------------------------------------

Public Enum ErrorType
    etSpelling = 1
    etGrammar = 2
    etMissingData = 3
    etCostAnomaly = 4
    etUnitError = 5
    etDescriptionError = 6
    etCalculationError = 7
    etFIDICError = 8
    etFormatError = 9
End Enum

Public Enum ErrorSeverity
    esInfo = 1
    esWarning = 2
    esCritical = 3
End Enum

'------------------------------------------------------------------------------
' Error Type to String Conversion
'------------------------------------------------------------------------------

Public Function ErrorTypeToString(ByVal errType As ErrorType) As String
    Select Case errType
        Case etSpelling: ErrorTypeToString = "Spelling"
        Case etGrammar: ErrorTypeToString = "Grammar"
        Case etMissingData: ErrorTypeToString = "Missing Data"
        Case etCostAnomaly: ErrorTypeToString = "Cost Anomaly"
        Case etUnitError: ErrorTypeToString = "Unit Error"
        Case etDescriptionError: ErrorTypeToString = "Description Error"
        Case etCalculationError: ErrorTypeToString = "Calculation Error"
        Case etFIDICError: ErrorTypeToString = "FIDIC Reference Error"
        Case etFormatError: ErrorTypeToString = "Format Error"
        Case Else: ErrorTypeToString = "Unknown"
    End Select
End Function

Public Function SeverityToString(ByVal sev As ErrorSeverity) As String
    Select Case sev
        Case esInfo: SeverityToString = "Info"
        Case esWarning: SeverityToString = "Warning"
        Case esCritical: SeverityToString = "Critical"
        Case Else: SeverityToString = "Unknown"
    End Select
End Function

'------------------------------------------------------------------------------
' Progress Indicator Helpers
'------------------------------------------------------------------------------

Public Sub ShowProgressMessage(ByVal message As String)
    ' This will be implemented with frmProgress
    ' For now, use status bar
    Application.StatusBar = message
End Sub

Public Sub ClearProgressMessage()
    Application.StatusBar = False
End Sub

'------------------------------------------------------------------------------
' Array Helpers
'------------------------------------------------------------------------------

' Check if array is empty
Public Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    On Error GoTo 0
End Function

' Get array length
Public Function ArrayLength(arr As Variant) As Long
    On Error Resume Next
    ArrayLength = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then ArrayLength = 0
    On Error GoTo 0
End Function
