# Developer Guide - Grammar & QS Checker Add-in

## Overview

This guide is for developers who want to understand, modify, or extend the Grammar & QS Checker Add-in.

## Architecture

### Technology Stack

- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Microsoft Excel 2016+
- **Format**: .xlam (Excel Add-in)
- **UI**: Custom Ribbon XML + UserForms

### Project Structure

```
Excel-Addin-Grammer/
├── src/
│   └── modules/              # VBA module files (.bas)
│       ├── ModUtility.bas    # Helper functions
│       ├── ModLogging.bas    # Change tracking & logging
│       ├── ModConfig.bas     # Configuration management
│       ├── ModSpelling.bas   # Spelling check engine
│       ├── ModGrammar.bas    # Grammar check engine
│       ├── ModMain.bas       # Main entry points
│       ├── ModQSValidator.bas         # QS orchestrator
│       ├── ModQSDictionary.bas        # QS terminology
│       ├── ModBOQAnalysis.bas         # BOQ structure analysis
│       ├── ModUnitValidator.bas       # Unit validation
│       ├── ModCostAnalysis.bas        # Cost/rate analysis
│       ├── ModDescriptionAnalysis.bas # Description checking
│       └── ModFIDIC.bas              # FIDIC validation
├── data/
│   ├── dictionaries/         # Spelling dictionaries (CSV)
│   ├── qs-data/             # QS reference data (CSV)
│   └── grammar-rules/       # Grammar rules (CSV)
├── ribbon/
│   └── customUI.xml         # Ribbon configuration
├── installer/
│   ├── install.bat          # Windows installer
│   └── install.sh           # Mac installer
└── docs/
    ├── user-guides/
    └── developer-guides/
```

## Module Dependencies

```
ModMain (Entry Point)
  ├── ModUI (User Interface)
  ├── ModSpelling (Spelling Check)
  │   └── ModUtility (Helpers)
  ├── ModGrammar (Grammar Check)
  │   └── ModUtility
  ├── ModQSValidator (QS Orchestrator)
  │   ├── ModQSDictionary
  │   ├── ModBOQAnalysis
  │   ├── ModUnitValidator
  │   ├── ModCostAnalysis
  │   ├── ModDescriptionAnalysis
  │   └── ModFIDIC
  ├── ModLogging (Change Tracking)
  └── ModConfig (Settings)
```

## Core Modules

### ModUtility.bas

**Purpose**: Common utility functions used across all modules

**Key Functions**:
- `NormalizeText()`: Text normalization
- `SplitIntoWords()`: Tokenization
- `LevenshteinDistance()`: Edit distance calculation for suggestions
- `GetWorksheetByName()`: Safe worksheet access
- Error type enumerations

**Usage**:
```vba
Dim cleanText As String
cleanText = ModUtility.NormalizeText(cellValue)

Dim distance As Integer
distance = ModUtility.LevenshteinDistance("hello", "helo") ' Returns 1
```

### ModLogging.bas

**Purpose**: Change tracking, undo functionality, and event logging

**Key Features**:
- Error record type definition
- Change stack for undo
- Log persistence to hidden worksheet
- Export functionality

**Data Structure**:
```vba
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
```

**Usage**:
```vba
' Log an event
Call ModLogging.LogEvent("Dictionary loaded", "INFO")

' Log a correction
Dim errRec As ModLogging.ErrorRecord
' ... populate errRec ...
Call ModLogging.LogCorrection(errRec)

' Undo last change
If ModLogging.UndoLastCorrection() Then
    MsgBox "Undone successfully"
End If
```

### ModConfig.bas

**Purpose**: Configuration and settings management

**Settings Storage**: Hidden worksheets `Settings` and `QS_Settings`

**Key Variables** (Module-level):
```vba
' Core Settings
Public EnableSpellingCheck As Boolean
Public EnableGrammarCheck As Boolean
Public DefaultLanguage As String

' QS Settings
Public EnableQSValidation As Boolean
Public EnableCostAnalysis As Boolean
Public CostAnomalyThresholdPercent As Double
```

**Usage**:
```vba
' Load settings on startup
Call ModConfig.LoadSettings

' Check if feature enabled
If ModConfig.EnableCostAnalysis Then
    ' Run cost analysis
End If

' Save settings after changes
ModConfig.CostAnomalyThresholdPercent = 60
Call ModConfig.SaveSettings
```

### ModSpelling.bas

**Purpose**: Spelling check engine with dictionary lookup

**Dictionary**: Stored in hidden worksheet `Dictionary_EN`, cached in memory

**Key Functions**:
- `InitializeDictionary()`: Load dictionary into memory
- `IsWordSpelledCorrectly()`: Check single word
- `CheckSpelling()`: Check text and return misspelled words
- `GetSpellingSuggestions()`: Generate suggestions using Levenshtein distance
- `AddWordToDictionary()`: Add custom words

**Algorithm**:
1. Load dictionary from worksheet into Collection (hash-like lookup)
2. Tokenize input text into words
3. Check each word against dictionary (O(1) lookup)
4. For misspelled words, calculate Levenshtein distance to all dictionary words
5. Return top N suggestions with smallest distance

**Usage**:
```vba
' Check if word is correct
If ModSpelling.IsWordSpelledCorrectly("concrete") Then
    ' Word is in dictionary
End If

' Get suggestions
Dim suggestions As Variant
suggestions = ModSpelling.GetSpellingSuggestions("concret")
' Returns array: ["concrete", "concert", ...]
```

### ModGrammar.bas

**Purpose**: Rule-based grammar checking

**Rules Storage**: Hidden worksheet `GrammarRules` or programmatic defaults

**Grammar Rule Structure**:
```vba
Public Type GrammarRule
    ruleID As String          ' e.g., "DOUBLE_SPACE"
    pattern As String         ' Text pattern to find
    replacement As String     ' Suggested correction
    Severity As ErrorSeverity
    Category As String
    Description As String
End Type
```

**Built-in Rules**:
- DOUBLE_SPACE: Multiple spaces → single space
- SPACE_BEFORE_PERIOD: " ." → "."
- NO_SPACE_AFTER_COMMA: "," → ", "
- And more...

**Usage**:
```vba
' Check grammar in text
Dim errors As Collection
Set errors = ModGrammar.CheckGrammar(cellText)

' Process each error
Dim i As Long
For i = 1 To errors.Count
    Dim errorInfo As Object
    Set errorInfo = errors(i)
    Debug.Print errorInfo("Description")
Next i
```

### ModMain.bas

**Purpose**: Main entry points and orchestration

**Ribbon Callbacks**:
- `CheckNow_Click()`: Main check button
- `CheckSpelling_Click()`: Spelling only
- `CheckGrammar_Click()`: Grammar only
- `CheckQS_Click()`: QS validation only
- `ShowSettings_Click()`: Open settings dialog
- `ShowHelp_Click()`: Display help

**Main Workflow**:
```
User clicks button
  ↓
Validate active workbook
  ↓
Get target range (selection or used range)
  ↓
Scan cells for errors
  ↓
Collect errors into g_ErrorCollection
  ↓
Display results dialog
  ↓
User accepts/rejects corrections
  ↓
Apply corrections & log changes
```

**Global Variables**:
```vba
Public g_ErrorCollection As Collection  ' Stores all found errors
```

## QS Modules

### ModQSValidator.bas

**Purpose**: Orchestrate all QS validation checks

**Main Function**:
```vba
Public Sub ScanRangeForQSErrors(ByRef targetRange As Range)
    ' Calls all sub-validators:
    ' - BOQ structure
    ' - Unit validation
    ' - Cost analysis
    ' - Description analysis
    ' - FIDIC validation (if enabled)
End Sub
```

### ModQSDictionary.bas

**Purpose**: Construction/QS terminology dictionary

**Data**: ~200 QS terms in hidden worksheet `QS_Dictionary`

**Features**:
- Term lookup
- Misspelling correction
- Standard unit suggestion for terms

### ModBOQAnalysis.bas

**Purpose**: BOQ structure and completeness validation

**Checks**:
- Missing quantities, rates, units, descriptions
- Calculation validation (Qty × Rate = Amount)
- Duplicate descriptions
- BOQ structure integrity

**Usage**:
```vba
Call ModBOQAnalysis.CheckMissingData(targetRange)
Call ModBOQAnalysis.ValidateCalculations(targetRange)
```

### ModUnitValidator.bas

**Purpose**: Unit of measurement validation

**Unit Masters**: Stored in `QS_UnitMasters` worksheet

**Common Units**: M, M², M³, NO, TONNE, LITRE, etc.

**Functions**:
- `IsValidUnit()`: Check if unit exists
- `ValidateUnits()`: Scan range for invalid units
- Suggest corrections for common misspellings

### ModCostAnalysis.bas

**Purpose**: Cost and rate validation with statistical analysis

**Checks**:
- Zero or negative rates
- Rate outliers using mean/standard deviation
- Configurable threshold for anomaly detection

**Algorithm**:
1. Collect all numeric values (potential rates)
2. Calculate mean and standard deviation
3. Flag values beyond threshold (default: 50% from mean)

### ModDescriptionAnalysis.bas

**Purpose**: BOQ description validation

**Checks**:
- Very short descriptions (likely incomplete)
- Duplicate descriptions (future: fuzzy matching)
- Missing required elements (future)

### ModFIDIC.bas

**Purpose**: FIDIC 1999 contract clause reference validation

**Data**: Complete FIDIC clause list in `QS_FIDICReferences`

**Features**:
- Extract clause numbers from text
- Validate against FIDIC 1999 database
- Suggest corrections

## Data Files

All data files are CSV format for easy editing and version control.

### english_dictionary.csv

**Columns**: Word, Length, Frequency, Category

**Size**: 150+ words (expand to 3,000-5,000 for production)

**Usage**: Import into hidden worksheet `Dictionary_EN`

### grammar_rules.csv

**Columns**: RuleID, Pattern, Replacement, Severity, Category, Description

**Usage**: Import into hidden worksheet `GrammarRules`

### qs_terminology.csv

**Columns**: Term, CorrectSpelling, CommonMisspellings, Category, Definition, StandardUnit, RegionalVariants, RelatedTerms

**Size**: 200+ construction/QS terms

**Usage**: Import into hidden worksheet `QS_Dictionary`

### unit_masters.csv

**Columns**: UnitCode, UnitName, ApplicableItems, ConversionFactor, Precision, CommonMisspellings

**Usage**: Import into hidden worksheet `QS_UnitMasters`

### fidic_references.csv

**Columns**: ClauseNumber, ClauseTitle, Requirements, RelatedClauses, PaymentRelated, TimelineRelated

**Size**: Complete FIDIC 1999 clause list (200+ clauses)

**Usage**: Import into hidden worksheet `QS_FIDICReferences`

## Ribbon Customization

### customUI.xml

Located in `ribbon/customUI.xml`

**Structure**:
- Custom tab: "Grammar & QS"
- Three groups: Spelling & Grammar, QS Validation, Help
- Buttons with callbacks to ModMain functions

**Adding New Button**:
```xml
<button id="btnMyNewFeature"
        label="My Feature"
        imageMso="SomeIcon"
        onAction="MyFeature_Click"
        screentip="My feature tooltip"/>
```

Then add callback in ModMain.bas:
```vba
Public Sub MyFeature_Click(control As IRibbonControl)
    MsgBox "My feature!"
End Sub
```

## Extending the Add-in

### Adding a New Grammar Rule

1. Open `data/grammar-rules/grammar_rules.csv`
2. Add new row:
   ```csv
   MY_RULE,"pattern","replacement",Warning,Category,Description
   ```
3. Re-import into `GrammarRules` worksheet
4. (Optional) Add special handling in `ModGrammar.FindPatternOccurrences()`

### Adding New QS Check

1. Create new function in appropriate module (e.g., `ModBOQAnalysis`)
2. Add call in `ModQSValidator.ScanRangeForQSErrors()`
3. Define error type in `ModUtility` if needed
4. Add configuration setting in `ModConfig` if needed

Example:
```vba
' In ModBOQAnalysis.bas
Public Sub CheckItemNumbering(ByRef targetRange As Range)
    ' Your validation logic
    ' Create ErrorRecord for issues
    ' Add to ModMain.g_ErrorCollection
End Sub

' In ModQSValidator.bas
Public Sub ScanRangeForQSErrors(...)
    ' ... existing checks ...
    Call ModBOQAnalysis.CheckItemNumbering(targetRange)
End Sub
```

### Adding QS Terminology

Edit `data/qs-data/qs_terminology.csv` and re-import to worksheet.

### Creating Custom Reports

Future enhancement: Add `ModQSReporting.bas` with functions like:
```vba
Public Sub GenerateQSReport(ByRef errorCollection As Collection)
    ' Create summary worksheet
    ' Export to PDF
End Sub
```

## Testing

### Unit Testing

VBA doesn't have built-in unit testing, but you can create test subroutines:

```vba
' In a test module
Public Sub TestLevenshteinDistance()
    Dim result As Integer
    result = ModUtility.LevenshteinDistance("kitten", "sitting")
    If result = 3 Then
        Debug.Print "PASS: LevenshteinDistance"
    Else
        Debug.Print "FAIL: Expected 3, got " & result
    End If
End Sub
```

### Integration Testing

1. Create test workbooks with known errors
2. Run add-in
3. Verify all errors are detected
4. Document test cases

### Performance Testing

Test on large datasets:
- 1,000 cells
- 10,000 cells
- 100,000 cells

Monitor execution time and optimize bottlenecks.

## Debugging

### Enable Debug Mode

In `ModMain.bas`:
```vba
Public Const DEBUG_MODE As Boolean = True

' In your code:
If DEBUG_MODE Then
    Debug.Print "Current value: " & someVariable
End If
```

### Immediate Window

Press `Ctrl+G` in VBA Editor to open Immediate Window for quick testing:
```vba
? ModSpelling.IsWordSpelledCorrectly("test")
? ModUtility.LevenshteinDistance("abc", "def")
```

### Breakpoints

Set breakpoints in VBA Editor (F9) to pause execution and inspect variables.

## Deployment

### Building the .xlam

See `ASSEMBLY_GUIDE.md` for detailed steps.

### Version Control

Export VBA modules regularly:
1. Right-click module in VBA Project Explorer
2. Export File
3. Save to `src/modules/`
4. Commit to Git

### Release Process

1. Increment version in `ModMain.ADDIN_VERSION`
2. Export all modules to `src/`
3. Test thoroughly
4. Build .xlam file
5. Create distribution package
6. Tag release in Git
7. Distribute

## Performance Optimization

### Bottlenecks

1. **Dictionary Lookup**: Cached in Collection for O(1) access
2. **Levenshtein Distance**: O(n×m) - limit comparisons to similar lengths
3. **Large Ranges**: Use batch processing, show progress

### Tips

- Disable screen updating: `Application.ScreenUpdating = False`
- Use arrays instead of cell-by-cell operations when possible
- Cache frequently accessed data
- Limit suggestion count (MAX_SUGGESTIONS = 5)

## Security Considerations

### Macro Security

- Code should be digitally signed for distribution
- Users must enable macros
- Don't execute arbitrary code from cells
- Validate all user inputs

### Data Privacy

- Add-in doesn't transmit data externally
- All processing is local
- Change log is stored locally in workbook

## Contributing

### Code Style

- Use meaningful variable names
- Comment complex logic
- Follow VBA naming conventions (PascalCase for functions, camelCase for variables)
- Add error handling with `On Error GoTo ErrorHandler`

### Pull Request Process

1. Fork repository
2. Create feature branch
3. Make changes and test
4. Update documentation
5. Submit PR with clear description

## Troubleshooting Common Issues

### "Compile Error: User-defined type not defined"

- Ensure all modules are imported in correct order
- ModUtility must be imported first (defines enums)

### "Object variable or With block variable not set"

- Check for `Set` keyword when assigning object variables
- Ensure worksheet/range exists before accessing

### "Type Mismatch"

- Check data types match function signatures
- Use `CStr()`, `CDbl()` for explicit conversions

## Future Enhancements

Potential additions:
- Real-time checking (Application events)
- Machine learning for better suggestions
- Multi-language support
- Custom user-defined rules
- Integration with external BOQ systems
- Advanced reporting (charts, dashboards)
- Collaboration features (comments, approvals)

---

**Version**: 1.0.0
**Last Updated**: 2025-12-12

For questions or contributions, see the project repository.
