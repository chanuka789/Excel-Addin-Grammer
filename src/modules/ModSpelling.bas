Attribute VB_Name = "ModSpelling"
'==============================================================================
' Module: ModSpelling
' Description: Spelling check engine with dictionary lookup
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

Private Const DICTIONARY_WORKSHEET_NAME As String = "Dictionary_EN"
Private Const MAX_SUGGESTIONS As Integer = 5
Private Const MAX_EDIT_DISTANCE As Integer = 2

' Dictionary cache
Private dictCache As Collection
Private dictLoaded As Boolean

'------------------------------------------------------------------------------
' Initialize Dictionary
'------------------------------------------------------------------------------

Public Sub InitializeDictionary()
    On Error GoTo ErrorHandler

    Set dictCache = New Collection
    dictLoaded = False

    ' Load dictionary from worksheet
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        MsgBox "Dictionary worksheet not found! Spelling check will not work.", vbCritical
        Exit Sub
    End If

    ' Load words into collection for fast lookup
    Dim lastRow As Long
    Dim i As Long
    Dim word As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow ' Skip header row
        word = UCase(Trim(ws.Cells(i, 1).Value))
        If Len(word) > 0 Then
            On Error Resume Next
            dictCache.Add word, word ' Use word as both item and key
            On Error GoTo ErrorHandler
        End If
    Next i

    dictLoaded = True
    Call ModLogging.LogEvent("Dictionary loaded: " & CStr(dictCache.Count) & " words", "INFO")

    Exit Sub

ErrorHandler:
    ' Continue loading even if duplicates exist
    If Err.Number <> 457 Then ' 457 = duplicate key
        MsgBox "Error loading dictionary: " & Err.Description, vbCritical
    End If
    Resume Next
End Sub

'------------------------------------------------------------------------------
' Check Single Word
'------------------------------------------------------------------------------

Public Function IsWordSpelledCorrectly(ByVal word As String) As Boolean
    ' Ensure dictionary is loaded
    If Not dictLoaded Then Call InitializeDictionary

    If dictCache Is Nothing Or dictCache.Count = 0 Then
        IsWordSpelledCorrectly = True ' Assume correct if no dictionary
        Exit Function
    End If

    ' Normalize word
    word = UCase(Trim(word))

    ' Skip very short words and numbers
    If Len(word) <= 1 Or ModUtility.IsNumericString(word) Then
        IsWordSpelledCorrectly = True
        Exit Function
    End If

    ' Check if word exists in dictionary
    On Error Resume Next
    Dim temp As Variant
    temp = dictCache(word)
    IsWordSpelledCorrectly = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Check Text and Return Misspelled Words
'------------------------------------------------------------------------------

Public Function CheckSpelling(ByVal inputText As String) As Collection
    ' Returns collection of misspelled words with their positions

    Set CheckSpelling = New Collection

    If Len(Trim(inputText)) = 0 Then Exit Function

    ' Split into words
    Dim words As Variant
    words = ModUtility.SplitIntoWords(inputText)

    If ModUtility.IsArrayEmpty(words) Then Exit Function

    ' Check each word
    Dim i As Long
    Dim word As String
    Dim errorInfo As Object

    For i = LBound(words) To UBound(words)
        word = Trim(words(i))

        If Len(word) > 0 Then
            If Not IsWordSpelledCorrectly(word) Then
                ' Create error information
                Set errorInfo = CreateObject("Scripting.Dictionary")
                errorInfo("Word") = word
                errorInfo("Position") = i

                On Error Resume Next
                CheckSpelling.Add errorInfo
                On Error GoTo 0
            End If
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Generate Spelling Suggestions
'------------------------------------------------------------------------------

Public Function GetSpellingSuggestions(ByVal misspelledWord As String) As Variant
    ' Returns array of suggested corrections

    Dim suggestions() As String
    Dim sugCount As Integer
    Dim maxSuggestions As Integer

    maxSuggestions = MAX_SUGGESTIONS
    sugCount = 0
    ReDim suggestions(0 To maxSuggestions - 1)

    If Not dictLoaded Then Call InitializeDictionary
    If dictCache Is Nothing Or dictCache.Count = 0 Then
        GetSpellingSuggestions = suggestions
        Exit Function
    End If

    misspelledWord = UCase(Trim(misspelledWord))

    ' Get dictionary worksheet for full scan
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If ws Is Nothing Then
        GetSpellingSuggestions = suggestions
        Exit Function
    End If

    ' Find closest matches using Levenshtein distance
    Dim lastRow As Long
    Dim i As Long
    Dim dictWord As String
    Dim distance As Integer

    ' Store words with their distances
    Dim wordDistances As Object
    Set wordDistances = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        dictWord = UCase(Trim(ws.Cells(i, 1).Value))

        If Len(dictWord) > 0 Then
            distance = ModUtility.LevenshteinDistance(misspelledWord, dictWord)

            ' Only consider words within reasonable edit distance
            If distance <= MAX_EDIT_DISTANCE Then
                If Not wordDistances.exists(dictWord) Then
                    wordDistances.Add dictWord, distance
                End If
            End If
        End If
    Next i

    ' Sort by distance and get top suggestions
    If wordDistances.Count > 0 Then
        Dim sortedWords As Variant
        sortedWords = GetSortedSuggestions(wordDistances, maxSuggestions)

        For i = 0 To UBound(sortedWords)
            If sugCount < maxSuggestions Then
                suggestions(sugCount) = sortedWords(i)
                sugCount = sugCount + 1
            End If
        Next i
    End If

    ' Resize array to actual count
    If sugCount > 0 Then
        ReDim Preserve suggestions(0 To sugCount - 1)
    Else
        ReDim suggestions(0 To 0)
        suggestions(0) = "(no suggestions)"
    End If

    GetSpellingSuggestions = suggestions
End Function

'------------------------------------------------------------------------------
' Sort Suggestions by Distance
'------------------------------------------------------------------------------

Private Function GetSortedSuggestions(ByRef wordDistances As Object, _
                                      ByVal maxCount As Integer) As Variant
    ' Simple bubble sort by distance value
    Dim words() As String
    Dim distances() As Integer
    Dim count As Integer
    Dim i As Long, j As Long
    Dim tempWord As String
    Dim tempDist As Integer

    count = wordDistances.Count
    If count > maxCount Then count = maxCount

    ReDim words(0 To wordDistances.Count - 1)
    ReDim distances(0 To wordDistances.Count - 1)

    ' Copy to arrays
    Dim keys As Variant
    keys = wordDistances.keys

    For i = 0 To wordDistances.Count - 1
        words(i) = keys(i)
        distances(i) = wordDistances(keys(i))
    Next i

    ' Bubble sort
    For i = 0 To UBound(words) - 1
        For j = i + 1 To UBound(words)
            If distances(j) < distances(i) Then
                ' Swap distances
                tempDist = distances(i)
                distances(i) = distances(j)
                distances(j) = tempDist

                ' Swap words
                tempWord = words(i)
                words(i) = words(j)
                words(j) = tempWord
            End If
        Next j
    Next i

    ' Return top N
    Dim result() As String
    ReDim result(0 To IIf(count - 1 < 0, 0, count - 1))

    For i = 0 To IIf(count - 1 < UBound(words), count - 1, UBound(words))
        result(i) = words(i)
    Next i

    GetSortedSuggestions = result
End Function

'------------------------------------------------------------------------------
' Add Word to Dictionary
'------------------------------------------------------------------------------

Public Sub AddWordToDictionary(ByVal newWord As String)
    On Error GoTo ErrorHandler

    newWord = Trim(newWord)
    If Len(newWord) = 0 Then Exit Sub

    ' Add to cache
    If Not dictLoaded Then Call InitializeDictionary

    Dim upperWord As String
    upperWord = UCase(newWord)

    On Error Resume Next
    dictCache.Add upperWord, upperWord
    On Error GoTo ErrorHandler

    ' Add to worksheet
    Dim ws As Worksheet
    Set ws = ModUtility.GetWorksheetByName(DICTIONARY_WORKSHEET_NAME, ThisWorkbook)

    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

        ws.Cells(lastRow, 1).Value = newWord
        ws.Cells(lastRow, 2).Value = Len(newWord)

        Call ModLogging.LogEvent("Added word to dictionary: " & newWord, "INFO")
    End If

    Exit Sub

ErrorHandler:
    If Err.Number <> 457 Then ' Ignore duplicate key errors
        MsgBox "Error adding word to dictionary: " & Err.Description, vbExclamation
    End If
End Sub

'------------------------------------------------------------------------------
' Check if Dictionary is Loaded
'------------------------------------------------------------------------------

Public Function IsDictionaryLoaded() As Boolean
    IsDictionaryLoaded = dictLoaded And Not (dictCache Is Nothing)
End Function

'------------------------------------------------------------------------------
' Get Dictionary Word Count
'------------------------------------------------------------------------------

Public Function GetDictionaryWordCount() As Long
    If dictCache Is Nothing Then
        GetDictionaryWordCount = 0
    Else
        GetDictionaryWordCount = dictCache.Count
    End If
End Function
