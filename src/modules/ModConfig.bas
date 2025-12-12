Attribute VB_Name = "ModConfig"
'==============================================================================
' Module: ModConfig
' Description: Configuration and settings management
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const SETTINGS_WORKSHEET_NAME As String = "Settings"
Private Const QS_SETTINGS_WORKSHEET_NAME As String = "QS_Settings"

'------------------------------------------------------------------------------
' Configuration Variables
'------------------------------------------------------------------------------

' Core Settings
Public EnableSpellingCheck As Boolean
Public EnableGrammarCheck As Boolean
Public EnableStyleCheck As Boolean
Public DefaultLanguage As String

' QS Settings
Public EnableQSValidation As Boolean
Public EnableBOQAnalysis As Boolean
Public EnableUnitValidation As Boolean
Public EnableCostAnalysis As Boolean
Public EnableFIDICValidation As Boolean
Public EnableIPCValidation As Boolean

' Thresholds
Public CostAnomalyThresholdPercent As Double
Public MinimumRateValue As Double
Public MaximumRateValue As Double

' UI Settings
Public AutoShowResults As Boolean
Public PlaySoundOnComplete As Boolean
Public ShowProgressBar As Boolean

'------------------------------------------------------------------------------
' Load Settings from Worksheet
'------------------------------------------------------------------------------

Public Sub LoadSettings()
    On Error Resume Next

    ' Set defaults first
    Call SetDefaultSettings

    ' Load from Settings worksheet if exists
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(SETTINGS_WORKSHEET_NAME, ThisWorkbook)

    If Not ws Is Nothing Then
        ' Read settings from worksheet
        EnableSpellingCheck = GetSettingValue(ws, "EnableSpellingCheck", True)
        EnableGrammarCheck = GetSettingValue(ws, "EnableGrammarCheck", True)
        EnableStyleCheck = GetSettingValue(ws, "EnableStyleCheck", False)
        DefaultLanguage = GetSettingValue(ws, "DefaultLanguage", "English")

        AutoShowResults = GetSettingValue(ws, "AutoShowResults", True)
        PlaySoundOnComplete = GetSettingValue(ws, "PlaySoundOnComplete", False)
        ShowProgressBar = GetSettingValue(ws, "ShowProgressBar", True)
    End If

    ' Load QS settings
    Dim qsWs As Worksheet
    Set qsWs = ModUtility.GetWorksheetByName(QS_SETTINGS_WORKSHEET_NAME, ThisWorkbook)

    If Not qsWs Is Nothing Then
        EnableQSValidation = GetSettingValue(qsWs, "EnableQSValidation", True)
        EnableBOQAnalysis = GetSettingValue(qsWs, "EnableBOQAnalysis", True)
        EnableUnitValidation = GetSettingValue(qsWs, "EnableUnitValidation", True)
        EnableCostAnalysis = GetSettingValue(qsWs, "EnableCostAnalysis", True)
        EnableFIDICValidation = GetSettingValue(qsWs, "EnableFIDICValidation", False)
        EnableIPCValidation = GetSettingValue(qsWs, "EnableIPCValidation", False)

        CostAnomalyThresholdPercent = GetSettingValue(qsWs, "CostAnomalyThreshold", 50)
        MinimumRateValue = GetSettingValue(qsWs, "MinimumRateValue", 0.01)
        MaximumRateValue = GetSettingValue(qsWs, "MaximumRateValue", 1000000)
    End If

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Save Settings to Worksheet
'------------------------------------------------------------------------------

Public Sub SaveSettings()
    On Error GoTo ErrorHandler

    ' Ensure worksheets exist
    Call EnsureSettingsWorksheets

    ' Save core settings
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SETTINGS_WORKSHEET_NAME)

    Call SetSettingValue(ws, "EnableSpellingCheck", EnableSpellingCheck)
    Call SetSettingValue(ws, "EnableGrammarCheck", EnableGrammarCheck)
    Call SetSettingValue(ws, "EnableStyleCheck", EnableStyleCheck)
    Call SetSettingValue(ws, "DefaultLanguage", DefaultLanguage)
    Call SetSettingValue(ws, "AutoShowResults", AutoShowResults)
    Call SetSettingValue(ws, "PlaySoundOnComplete", PlaySoundOnComplete)
    Call SetSettingValue(ws, "ShowProgressBar", ShowProgressBar)

    ' Save QS settings
    Dim qsWs As Worksheet
    Set qsWs = ThisWorkbook.Worksheets(QS_SETTINGS_WORKSHEET_NAME)

    Call SetSettingValue(qsWs, "EnableQSValidation", EnableQSValidation)
    Call SetSettingValue(qsWs, "EnableBOQAnalysis", EnableBOQAnalysis)
    Call SetSettingValue(qsWs, "EnableUnitValidation", EnableUnitValidation)
    Call SetSettingValue(qsWs, "EnableCostAnalysis", EnableCostAnalysis)
    Call SetSettingValue(qsWs, "EnableFIDICValidation", EnableFIDICValidation)
    Call SetSettingValue(qsWs, "EnableIPCValidation", EnableIPCValidation)
    Call SetSettingValue(qsWs, "CostAnomalyThreshold", CostAnomalyThresholdPercent)
    Call SetSettingValue(qsWs, "MinimumRateValue", MinimumRateValue)
    Call SetSettingValue(qsWs, "MaximumRateValue", MaximumRateValue)

    Exit Sub

ErrorHandler:
    MsgBox "Error saving settings: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Default Settings
'------------------------------------------------------------------------------

Public Sub SetDefaultSettings()
    ' Core defaults
    EnableSpellingCheck = True
    EnableGrammarCheck = True
    EnableStyleCheck = False
    DefaultLanguage = "English"

    ' QS defaults
    EnableQSValidation = True
    EnableBOQAnalysis = True
    EnableUnitValidation = True
    EnableCostAnalysis = True
    EnableFIDICValidation = False
    EnableIPCValidation = False

    ' Thresholds
    CostAnomalyThresholdPercent = 50
    MinimumRateValue = 0.01
    MaximumRateValue = 1000000

    ' UI defaults
    AutoShowResults = True
    PlaySoundOnComplete = False
    ShowProgressBar = True
End Sub

'------------------------------------------------------------------------------
' Create Default Settings (First Time Setup)
'------------------------------------------------------------------------------

Public Sub CreateDefaultSettings()
    Call SetDefaultSettings
    Call EnsureSettingsWorksheets
    Call SaveSettings

    Call ModLogging.LogEvent("Default settings created", "INFO")
End Sub

'------------------------------------------------------------------------------
' Helper Functions
'------------------------------------------------------------------------------

Private Function GetSettingValue(ByRef ws As Worksheet, _
                                 ByVal settingName As String, _
                                 ByVal defaultValue As Variant) As Variant
    On Error Resume Next

    Dim lastRow As Long
    Dim i As Long
    Dim foundValue As Variant

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Search for setting name in column A
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = settingName Then
            foundValue = ws.Cells(i, 2).Value
            If Not IsEmpty(foundValue) Then
                GetSettingValue = foundValue
                Exit Function
            End If
        End If
    Next i

    ' Not found, return default
    GetSettingValue = defaultValue

    On Error GoTo 0
End Function

Private Sub SetSettingValue(ByRef ws As Worksheet, _
                            ByVal settingName As String, _
                            ByVal value As Variant)
    On Error Resume Next

    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Search for existing setting
    found = False
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = settingName Then
            ws.Cells(i, 2).Value = value
            found = True
            Exit For
        End If
    Next i

    ' If not found, add new row
    If Not found Then
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = settingName
        ws.Cells(lastRow, 2).Value = value
    End If

    On Error GoTo 0
End Sub

Private Sub EnsureSettingsWorksheets()
    ' Ensure Settings worksheet exists
    Call EnsureSettingsWorksheet(SETTINGS_WORKSHEET_NAME)
    Call EnsureSettingsWorksheet(QS_SETTINGS_WORKSHEET_NAME)
End Sub

Private Sub EnsureSettingsWorksheet(ByVal wsName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new settings worksheet
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName

        ' Add headers
        ws.Cells(1, 1).Value = "SettingName"
        ws.Cells(1, 2).Value = "SettingValue"
        ws.Cells(1, 3).Value = "SettingType"
        ws.Cells(1, 4).Value = "Description"

        ' Format headers
        ws.Range("A1:D1").Font.Bold = True
        ws.Range("A1:D1").Interior.Color = RGB(200, 200, 200)
        ws.Columns("A:D").AutoFit

        ' Hide the worksheet
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

'------------------------------------------------------------------------------
' Reset to Defaults
'------------------------------------------------------------------------------

Public Sub ResetToDefaults()
    Dim result As VbMsgBoxResult

    result = MsgBox("This will reset all settings to default values. Continue?", _
                    vbYesNo + vbQuestion, "Reset Settings")

    If result = vbYes Then
        Call SetDefaultSettings
        Call SaveSettings
        MsgBox "Settings have been reset to defaults.", vbInformation
    End If
End Sub
