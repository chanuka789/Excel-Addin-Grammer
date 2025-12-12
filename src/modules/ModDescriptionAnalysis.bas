Attribute VB_Name = "ModDescriptionAnalysis"
'==============================================================================
' Module: ModDescriptionAnalysis
' Description: BOQ description standardization and analysis
' Author: Grammar & QS Add-in Development Team
' Version: 1.0.0
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Analyze Descriptions in Range
'------------------------------------------------------------------------------

Public Sub AnalyzeDescriptions(ByRef targetRange As Range)
    Call CheckIncompleteDescriptions(targetRange)
End Sub

'------------------------------------------------------------------------------
' Check for Incomplete Descriptions
'------------------------------------------------------------------------------

Private Sub CheckIncompleteDescriptions(ByRef targetRange As Range)
    Dim cell As Range
    Dim desc As String
    Dim errRecord As ModLogging.ErrorRecord

    For Each cell In targetRange.Cells
        desc = Trim(cell.Value)

        ' Check text cells
        If Len(desc) > 0 And Not IsNumeric(desc) Then
            ' Very short descriptions (likely incomplete)
            If Len(desc) < 10 And Len(desc) > 2 Then
                ' Might be incomplete
                errRecord.CellAddress = cell.Address
                errRecord.SheetName = cell.Worksheet.Name
                errRecord.WorkbookName = cell.Worksheet.Parent.Name
                errRecord.errorType = ModUtility.etDescriptionError
                errRecord.OriginalText = desc
                errRecord.CorrectedText = "(Review - Description may be incomplete)"
                errRecord.Severity = ModUtility.esInfo
                errRecord.Category = "Short Description"
                errRecord.Timestamp = ModUtility.GetTimestamp()
                errRecord.Applied = False

                ModMain.g_ErrorCollection.Add errRecord
            End If
        End If
    Next cell
End Sub
