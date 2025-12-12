Attribute VB_Name = "ModGrammar"
'==============================================================================
' Module: ModGrammar
' Description: Grammar checking engine with rule-based validation
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const GRAMMAR_RULES_WORKSHEET_NAME As String = "GrammarRules"

' Grammar rules cache
Private grammarRules As Collection
Private rulesLoaded As Boolean

'------------------------------------------------------------------------------
' Grammar Rule Type
'------------------------------------------------------------------------------

Public Type GrammarRule
    ruleID As String
    pattern As String
    replacement As String
    Severity As ModUtility.ErrorSeverity
    Category As String
    Description As String
End Type

'------------------------------------------------------------------------------
' Initialize Grammar Rules
'------------------------------------------------------------------------------

Public Sub InitializeGrammarRules()
    On Error GoTo ErrorHandler

    Set grammarRules = New Collection
    rulesLoaded = False

    ' Load rules from worksheet
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(GRAMMAR_RULES_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        ' Create default rules programmatically
        Call CreateDefaultGrammarRules
        Exit Sub
    End If

    ' Load rules from worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rule As GrammarRule

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow ' Skip header
        rule.ruleID = Trim(ws.Cells(i, 1).Value)
        rule.pattern = Trim(ws.Cells(i, 2).Value)
        rule.replacement = Trim(ws.Cells(i, 3).Value)
        rule.Severity = ParseSeverity(ws.Cells(i, 4).Value)
        rule.Category = Trim(ws.Cells(i, 5).Value)
        rule.Description = Trim(ws.Cells(i, 6).Value)

        If Len(rule.ruleID) > 0 Then
            grammarRules.Add rule
        End If
    Next i

    rulesLoaded = True
    Call ModLogging.LogEvent("Grammar rules loaded: " & CStr(grammarRules.Count) & " rules", "INFO")

    Exit Sub

ErrorHandler:
    MsgBox "Error loading grammar rules: " & Err.Description, vbCritical
    Call CreateDefaultGrammarRules
End Sub

'------------------------------------------------------------------------------
' Create Default Grammar Rules (if worksheet not found)
'------------------------------------------------------------------------------

Private Sub CreateDefaultGrammarRules()
    Set grammarRules = New Collection

    Dim rule As GrammarRule

    ' Rule 1: Double spaces
    rule.ruleID = "DOUBLE_SPACE"
    rule.pattern = "  " ' Two spaces
    rule.replacement = " " ' Single space
    rule.Severity = ModUtility.esWarning
    rule.Category = "Spacing"
    rule.Description = "Multiple consecutive spaces"
    grammarRules.Add rule

    ' Rule 2: Space before punctuation
    rule.ruleID = "SPACE_BEFORE_PERIOD"
    rule.pattern = " ."
    rule.replacement = "."
    rule.Severity = ModUtility.esWarning
    rule.Category = "Punctuation"
    rule.Description = "Space before period"
    grammarRules.Add rule

    ' Rule 3: Space before comma
    rule.ruleID = "SPACE_BEFORE_COMMA"
    rule.pattern = " ,"
    rule.replacement = ","
    rule.Severity = ModUtility.esWarning
    rule.Category = "Punctuation"
    rule.Description = "Space before comma"
    grammarRules.Add rule

    ' Rule 4: No space after period
    rule.ruleID = "NO_SPACE_AFTER_PERIOD"
    rule.pattern = "."
    rule.replacement = ". "
    rule.Severity = ModUtility.esWarning
    rule.Category = "Punctuation"
    rule.Description = "Missing space after period"
    grammarRules.Add rule

    ' Rule 5: No space after comma
    rule.ruleID = "NO_SPACE_AFTER_COMMA"
    rule.pattern = ","
    rule.replacement = ", "
    rule.Severity = ModUtility.esWarning
    rule.Category = "Punctuation"
    rule.Description = "Missing space after comma"
    grammarRules.Add rule

    rulesLoaded = True
End Sub

'------------------------------------------------------------------------------
' Check Grammar in Text
'------------------------------------------------------------------------------

Public Function CheckGrammar(ByVal inputText As String) As Collection
    ' Returns collection of grammar errors found

    Set CheckGrammar = New Collection

    If Len(Trim(inputText)) = 0 Then Exit Function

    ' Ensure rules are loaded
    If Not rulesLoaded Then Call InitializeGrammarRules

    ' Apply each rule
    Dim rule As GrammarRule
    Dim i As Long
    Dim errorInfo As Object

    For i = 1 To grammarRules.Count
        rule = grammarRules(i)

        ' Check if pattern exists in text
        Dim errors As Collection
        Set errors = FindPatternOccurrences(inputText, rule)

        ' Add all occurrences to results
        Dim j As Long
        For j = 1 To errors.Count
            CheckGrammar.Add errors(j)
        Next j
    Next i
End Function

'------------------------------------------------------------------------------
' Find All Occurrences of Pattern
'------------------------------------------------------------------------------

Private Function FindPatternOccurrences(ByVal inputText As String, _
                                        ByRef rule As GrammarRule) As Collection
    Set FindPatternOccurrences = New Collection

    Dim pos As Long
    Dim startPos As Long
    Dim errorInfo As Object

    startPos = 1

    ' Special handling for certain patterns
    Select Case rule.ruleID
        Case "DOUBLE_SPACE"
            ' Find all double spaces
            Do
                pos = InStr(startPos, inputText, rule.pattern)
                If pos > 0 Then
                    Set errorInfo = CreateObject("Scripting.Dictionary")
                    errorInfo("RuleID") = rule.ruleID
                    errorInfo("Pattern") = rule.pattern
                    errorInfo("Replacement") = rule.replacement
                    errorInfo("Position") = pos
                    errorInfo("Length") = Len(rule.pattern)
                    errorInfo("Severity") = rule.Severity
                    errorInfo("Category") = rule.Category
                    errorInfo("Description") = rule.Description

                    FindPatternOccurrences.Add errorInfo
                    startPos = pos + Len(rule.pattern)
                Else
                    Exit Do
                End If
            Loop

        Case "NO_SPACE_AFTER_PERIOD", "NO_SPACE_AFTER_COMMA"
            ' Check for period/comma followed immediately by letter
            Dim char As String
            Dim punctuation As String

            If rule.ruleID = "NO_SPACE_AFTER_PERIOD" Then
                punctuation = "."
            Else
                punctuation = ","
            End If

            For pos = 1 To Len(inputText) - 1
                If Mid(inputText, pos, 1) = punctuation Then
                    char = Mid(inputText, pos + 1, 1)
                    ' Check if next char is a letter (not space, not punctuation)
                    If char Like "[A-Za-z]" Then
                        Set errorInfo = CreateObject("Scripting.Dictionary")
                        errorInfo("RuleID") = rule.ruleID
                        errorInfo("Pattern") = punctuation
                        errorInfo("Replacement") = punctuation & " "
                        errorInfo("Position") = pos
                        errorInfo("Length") = 1
                        errorInfo("Severity") = rule.Severity
                        errorInfo("Category") = rule.Category
                        errorInfo("Description") = rule.Description

                        FindPatternOccurrences.Add errorInfo
                    End If
                End If
            Next pos

        Case Else
            ' Generic pattern matching
            Do
                pos = InStr(startPos, inputText, rule.pattern)
                If pos > 0 Then
                    Set errorInfo = CreateObject("Scripting.Dictionary")
                    errorInfo("RuleID") = rule.ruleID
                    errorInfo("Pattern") = rule.pattern
                    errorInfo("Replacement") = rule.replacement
                    errorInfo("Position") = pos
                    errorInfo("Length") = Len(rule.pattern)
                    errorInfo("Severity") = rule.Severity
                    errorInfo("Category") = rule.Category
                    errorInfo("Description") = rule.Description

                    FindPatternOccurrences.Add errorInfo
                    startPos = pos + Len(rule.pattern)
                Else
                    Exit Do
                End If
            Loop
    End Select
End Function

'------------------------------------------------------------------------------
' Apply Grammar Correction
'------------------------------------------------------------------------------

Public Function ApplyGrammarCorrection(ByVal inputText As String, _
                                       ByRef errorInfo As Object) As String
    Dim pattern As String
    Dim replacement As String
    Dim pos As Long

    pattern = errorInfo("Pattern")
    replacement = errorInfo("Replacement")
    pos = errorInfo("Position")

    ' Replace the error at the specific position
    Dim beforeText As String
    Dim afterText As String

    beforeText = Left(inputText, pos - 1)
    afterText = Mid(inputText, pos + Len(pattern))

    ApplyGrammarCorrection = beforeText & replacement & afterText
End Function

'------------------------------------------------------------------------------
' Get Grammar Suggestion
'------------------------------------------------------------------------------

Public Function GetGrammarSuggestion(ByRef errorInfo As Object) As String
    GetGrammarSuggestion = errorInfo("Replacement")
End Function

'------------------------------------------------------------------------------
' Add Custom Grammar Rule
'------------------------------------------------------------------------------

Public Sub AddGrammarRule(ByVal ruleID As String, _
                         ByVal pattern As String, _
                         ByVal replacement As String, _
                         ByVal Severity As ModUtility.ErrorSeverity, _
                         ByVal Category As String, _
                         ByVal Description As String)
    Dim rule As GrammarRule

    rule.ruleID = ruleID
    rule.pattern = pattern
    rule.replacement = replacement
    rule.Severity = Severity
    rule.Category = Category
    rule.Description = Description

    If Not rulesLoaded Then Call InitializeGrammarRules

    grammarRules.Add rule

    ' Optionally save to worksheet
    Call SaveRuleToWorksheet(rule)
End Sub

'------------------------------------------------------------------------------
' Save Rule to Worksheet
'------------------------------------------------------------------------------

Private Sub SaveRuleToWorksheet(ByRef rule As GrammarRule)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(GRAMMAR_RULES_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = rule.ruleID
    ws.Cells(lastRow, 2).Value = rule.pattern
    ws.Cells(lastRow, 3).Value = rule.replacement
    ws.Cells(lastRow, 4).Value = ModUtility.SeverityToString(rule.Severity)
    ws.Cells(lastRow, 5).Value = rule.Category
    ws.Cells(lastRow, 6).Value = rule.Description

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Parse Severity String
'------------------------------------------------------------------------------

Private Function ParseSeverity(ByVal sevString As String) As ModUtility.ErrorSeverity
    Select Case UCase(Trim(sevString))
        Case "INFO"
            ParseSeverity = ModUtility.esInfo
        Case "WARNING"
            ParseSeverity = ModUtility.esWarning
        Case "CRITICAL"
            ParseSeverity = ModUtility.esCritical
        Case Else
            ParseSeverity = ModUtility.esWarning
    End Select
End Function

'------------------------------------------------------------------------------
' Get Grammar Rules Count
'------------------------------------------------------------------------------

Public Function GetGrammarRulesCount() As Long
    If grammarRules Is Nothing Then
        GetGrammarRulesCount = 0
    Else
        GetGrammarRulesCount = grammarRules.Count
    End If
End Function
