Attribute VB_Name = "ModUnitValidator"
'==============================================================================
' Module: ModUnitValidator
' Description: Unit validation and standardization
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const UNIT_MASTERS_WORKSHEET_NAME As String = "QS_UnitMasters"

Private unitMasters As Collection
Private unitsLoaded As Boolean

'------------------------------------------------------------------------------
' Initialize Unit Validator
'------------------------------------------------------------------------------

Public Sub InitializeUnitValidator()
    On Error GoTo ErrorHandler

    Set unitMasters = New Collection
    unitsLoaded = False

    ' Load from worksheet
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(UNIT_MASTERS_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        ' Create default units
        Call CreateDefaultUnits
        Exit Sub
    End If

    ' Load units
    Dim lastRow As Long
    Dim i As Long
    Dim unitCode As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        unitCode = UCase(Trim(ws.Cells(i, 1).Value))
        If Len(unitCode) > 0 Then
            On Error Resume Next
            unitMasters.Add unitCode, unitCode
            On Error GoTo ErrorHandler
        End If
    Next i

    unitsLoaded = True
    Call ModLogging.LogEvent("Unit masters loaded: " & CStr(unitMasters.Count) & " units", "INFO")

    Exit Sub

ErrorHandler:
    If Err.Number <> 457 Then
        Call ModLogging.LogEvent("Error loading unit masters: " & Err.Description, "ERROR")
    End If
    Resume Next
End Sub

'------------------------------------------------------------------------------
' Create Default Units
'------------------------------------------------------------------------------

Private Sub CreateDefaultUnits()
    Set unitMasters = New Collection

    ' Common construction units
    Dim commonUnits As Variant
    commonUnits = Array("M", "M²", "M³", "M2", "M3", "NO", "NO.", _
                       "KG", "TONNE", "LITRE", "L", "SQ.FT", "CU.FT", _
                       "SQFT", "CUFT", "SQM", "CUM", "RM", "LM", "SM")

    Dim i As Long
    For i = LBound(commonUnits) To UBound(commonUnits)
        On Error Resume Next
        unitMasters.Add UCase(commonUnits(i)), UCase(commonUnits(i))
        On Error GoTo 0
    Next i

    unitsLoaded = True
End Sub

'------------------------------------------------------------------------------
' Validate Units in Range
'------------------------------------------------------------------------------

Public Sub ValidateUnits(ByRef targetRange As Range)
    If Not unitsLoaded Then Call InitializeUnitValidator

    Dim cell As Range
    Dim unitText As String
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        unitText = UCase(Trim(cell.Value))

        ' Check if cell might be a unit cell (short text, not numeric)
        If Len(unitText) > 0 And Len(unitText) <= 10 And Not IsNumeric(unitText) Then
            If Not IsValidUnit(unitText) Then
                ' Get suggestion
                Dim suggestion As String
                suggestion = GetUnitSuggestion(unitText)

                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etUnitError
                errRecord.OriginalText = cell.Value
                errRecord.CorrectedText = suggestion
                errRecord.Severity = ModUtility.esWarning
                errRecord.Category = "Invalid Unit"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            End If
        End If
    Next cell
End Sub

'------------------------------------------------------------------------------
' Check if Unit is Valid
'------------------------------------------------------------------------------

Public Function IsValidUnit(ByVal unitCode As String) As Boolean
    If Not unitsLoaded Then Call InitializeUnitValidator

    unitCode = UCase(Trim(unitCode))

    On Error Resume Next
    Dim temp As Variant
    temp = unitMasters(unitCode)
    IsValidUnit = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Get Unit Suggestion
'------------------------------------------------------------------------------

Private Function GetUnitSuggestion(ByVal invalidUnit As String) As String
    ' Simple suggestions based on common mistakes

    Select Case UCase(invalidUnit)
        Case "MM3", "MM³": GetUnitSuggestion = "M³"
        Case "MM2", "MM²": GetUnitSuggestion = "M²"
        Case "METER", "METRES": GetUnitSuggestion = "M"
        Case "SQM", "SQ.M": GetUnitSuggestion = "M²"
        Case "CUM", "CU.M": GetUnitSuggestion = "M³"
        Case "NUMBER", "NOS": GetUnitSuggestion = "NO"
        Case "TON": GetUnitSuggestion = "TONNE"
        Case Else: GetUnitSuggestion = "(verify unit)"
    End Select
End Function
