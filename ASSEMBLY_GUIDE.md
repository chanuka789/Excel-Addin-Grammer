# Excel Add-in Assembly Guide

This guide provides step-by-step instructions to build the **GrammarChecker_QS.xlam** Excel Add-in from source code.

## Prerequisites

- Microsoft Excel 2016 or later (Windows or Mac)
- Basic understanding of VBA and Excel Add-ins
- Text editor for viewing source files

## Assembly Steps

### Step 1: Create New Excel Add-in File

1. Open Microsoft Excel
2. Create a new blank workbook
3. Press `Alt + F11` (Windows) or `Option + F11` (Mac) to open VBA Editor
4. In the VBA Editor, go to **Tools** > **VBAProject Properties**
5. Set **Project Name**: `GrammarChecker_QS_AddIn`
6. Click OK

### Step 2: Import VBA Modules

1. In the VBA Editor, right-click on **VBAProject (Book1)**
2. Select **Import File...**
3. Import all modules from `src/modules/` in this order:

**Core Modules (Import First):**
- `ModUtility.bas` (Helper functions - must be first)
- `ModLogging.bas` (Logging functionality)
- `ModConfig.bas` (Configuration settings)
- `ModSpelling.bas` (Spelling check engine)
- `ModGrammar.bas` (Grammar check engine)
- `ModMain.bas` (Main entry points)
- `ModUI.bas` (UI coordination)

**QS Modules (Import After Core):**
- `ModQSDictionary.bas` (QS terminology)
- `ModBOQAnalysis.bas` (BOQ structure analysis)
- `ModUnitValidator.bas` (Unit validation)
- `ModCostAnalysis.bas` (Cost/rate validation)
- `ModDescriptionAnalysis.bas` (Description standardization)
- `ModFIDIC.bas` (FIDIC clause validation)
- `ModQSValidator.bas` (QS orchestrator)
- `ModQSReporting.bas` (QS report generator)

### Step 3: Import UserForms

1. In the VBA Editor, right-click on **VBAProject**
2. Select **Import File...**
3. Import all forms from `src/forms/`:

- `frmResults.frm` (Results dialog)
- `frmSettings.frm` (Settings dialog)
- `frmProgress.frm` (Progress indicator)
- `frmQSSettings.frm` (QS-specific settings)
- `frmQSReport.frm` (QS report viewer)

**Note**: Each .frm file should have a corresponding .frx file (binary form data). Import the .frm files; Excel will automatically load the .frx files.

### Step 4: Import Class Modules (if any)

1. Import any .cls files from `src/classes/` if present
2. Currently no class modules are required for v1.0

### Step 5: Create Hidden Data Worksheets

1. Close the VBA Editor (return to Excel)
2. Create the following worksheets (they will be hidden later):

**Core Data Worksheets:**
- `Dictionary_EN` - Spelling dictionary
- `GrammarRules` - Grammar validation rules
- `Settings` - User settings and preferences

**QS Data Worksheets:**
- `QS_Dictionary` - QS terminology database
- `QS_UnitMasters` - Standard construction units
- `QS_FIDICReferences` - FIDIC 1999 clause database
- `QS_ItemTemplates` - BOQ description templates
- `QS_DescriptionPatterns` - Description validation patterns
- `QS_Settings` - QS module configuration

### Step 6: Import Data into Worksheets

**For Dictionary_EN:**
1. Open `data/dictionaries/english_dictionary.csv`
2. Copy all data
3. Paste into `Dictionary_EN` worksheet starting at cell A1
4. Column headers: `Word`, `Length`, `Frequency`, `Category`

**For GrammarRules:**
1. Open `data/grammar-rules/grammar_rules.csv`
2. Copy and paste into `GrammarRules` worksheet
3. Column headers: `RuleID`, `Pattern`, `Replacement`, `Severity`, `Category`, `Description`

**For QS_Dictionary:**
1. Open `data/qs-data/qs_terminology.csv`
2. Copy and paste into `QS_Dictionary` worksheet
3. Column headers: `Term`, `CorrectSpelling`, `CommonMisspellings`, `Category`, `Definition`, `StandardUnit`, `RegionalVariants`, `RelatedTerms`

**For QS_UnitMasters:**
1. Open `data/qs-data/unit_masters.csv`
2. Copy and paste into `QS_UnitMasters` worksheet
3. Column headers: `UnitCode`, `UnitName`, `ApplicableItems`, `ConversionFactor`, `Precision`, `CommonMisspellings`

**For QS_FIDICReferences:**
1. Open `data/qs-data/fidic_references.csv`
2. Copy and paste into `QS_FIDICReferences` worksheet
3. Column headers: `ClauseNumber`, `ClauseTitle`, `Requirements`, `RelatedClauses`, `PaymentRelated`, `TimelineRelated`

**For QS_ItemTemplates:**
1. Open `data/qs-data/item_templates.csv`
2. Copy and paste into `QS_ItemTemplates` worksheet
3. Column headers: `ItemType`, `TemplateFormat`, `RequiredElements`, `Example`

**For QS_DescriptionPatterns:**
1. Open `data/qs-data/description_patterns.csv`
2. Copy and paste into `QS_DescriptionPatterns` worksheet
3. Column headers: `PatternID`, `ItemCategory`, `RegexPattern`, `RequiredFields`, `ValidationLevel`

**For Settings and QS_Settings:**
These will be populated at runtime by the add-in. Leave them with just column headers:

**Settings headers**: `SettingName`, `SettingValue`, `SettingType`, `Description`

**QS_Settings headers**: `SettingName`, `SettingValue`, `SettingType`, `Description`

### Step 7: Hide Data Worksheets

1. Right-click each data worksheet tab (all except Sheet1 if you have it)
2. Select **Hide**
3. Or better yet, use VBA to set them as `xlSheetVeryHidden`:

```vba
' Run this in Immediate Window (Ctrl+G in VBA Editor):
ThisWorkbook.Worksheets("Dictionary_EN").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("GrammarRules").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("Settings").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_Dictionary").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_UnitMasters").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_FIDICReferences").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_ItemTemplates").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_DescriptionPatterns").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("Settings").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_Settings").Visible = xlSheetVeryHidden
```

### Step 8: Configure ThisWorkbook Module

1. In VBA Editor, double-click **ThisWorkbook** in the project tree
2. Add the following code:

```vba
Private Sub Workbook_Open()
    ' This runs when the add-in loads
    Call InitializeAddIn
End Sub

Private Sub Workbook_AddinInstall()
    ' This runs when add-in is first installed
    Call OnInstall
End Sub

Private Sub Workbook_AddinUninstall()
    ' This runs when add-in is uninstalled
    Call OnUninstall
End Sub

Private Sub InitializeAddIn()
    ' Load settings from Settings worksheet
    Call ModConfig.LoadSettings

    ' Initialize QS module if enabled
    Call ModQSValidator.InitializeQS

    ' Log startup
    Call ModLogging.LogEvent("Add-in loaded successfully", "INFO")
End Sub

Private Sub OnInstall()
    ' First-time setup
    Call ModConfig.CreateDefaultSettings
    MsgBox "Grammar & QS Add-in installed successfully!" & vbCrLf & _
           "Look for 'Spelling & Grammar' buttons in the ribbon.", _
           vbInformation, "Installation Complete"
End Sub

Private Sub OnUninstall()
    ' Cleanup if needed
    MsgBox "Grammar & QS Add-in has been uninstalled.", vbInformation
End Sub
```

### Step 9: Configure Custom Ribbon (Advanced)

**Note**: Custom ribbon configuration requires Office 2007+ and is advanced. For simplicity in v1.0, you can skip this and use the default Add-ins tab.

**If you want custom ribbon:**

1. Close Excel
2. Rename your .xlam file to .zip
3. Extract the contents
4. Navigate to `customUI/` folder (create if it doesn't exist)
5. Copy `ribbon/customUI.xml` from this project into that folder
6. Update `_rels/.rels` to reference customUI/customUI.xml
7. Re-zip the contents and rename back to .xlam

**Simpler Alternative**: Use the built-in Add-ins tab and create a toolbar using VBA:

Add this to `ModMain.bas`:

```vba
Sub CreateToolbar()
    ' Creates simple toolbar buttons (Excel 2003 style, works in newer versions)
    ' This is simpler than ribbon customization
    ' Users can access via Add-ins tab
End Sub
```

### Step 10: Save as Excel Add-in (.xlam)

1. Go to **File** > **Save As**
2. Choose save location (your working directory for now)
3. **Save as type**: Select **Excel Add-in (*.xlam)**
4. **File name**: `GrammarChecker_QS.xlam`
5. Click **Save**

### Step 11: Test the Add-in

1. Close the .xlam file
2. Go to **File** > **Options** > **Add-ins**
3. At the bottom, select **Manage: Excel Add-ins** > Click **Go...**
4. Click **Browse...** and select your `GrammarChecker_QS.xlam` file
5. Check the box next to "GrammarChecker_QS" to enable it
6. Click OK
7. Open a new workbook and test:
   - Type some text with spelling errors
   - Run the add-in (via Add-ins tab or custom ribbon)
   - Verify functionality

### Step 12: Enable Macros and Trust

**For testing:**
1. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
2. **Macro Settings**: Select "Enable all macros" (for development only)
3. **Trusted Locations**: Add the folder containing your .xlam file

**For distribution:**
1. Digitally sign the .xlam file with a code signing certificate
2. Users will need to trust your certificate or lower macro security

### Step 13: Final Distribution Package

Create a distribution folder with:

```
GrammarChecker_QS_v1.0/
├── GrammarChecker_QS.xlam
├── install.bat (Windows installer)
├── install.sh (Mac installer)
├── README.txt (Installation instructions)
├── User_Guide.pdf
└── QS_Features_Guide.pdf
```

## Troubleshooting

**Problem**: Add-in doesn't load
- **Solution**: Check Trust Center settings, enable macros

**Problem**: Compile errors when opening
- **Solution**: Ensure all modules are imported in correct order (dependencies)

**Problem**: Data worksheets not found errors
- **Solution**: Verify all worksheet names match exactly (case-sensitive)

**Problem**: Forms don't display correctly
- **Solution**: Ensure both .frm and .frx files are in same directory during import

**Problem**: Ribbon buttons don't appear
- **Solution**: Ribbon customization is complex; use Add-ins tab as alternative

## Version Control

To maintain this project in version control:
- Export all VBA modules periodically: Right-click module > Export File
- Keep exported .bas, .frm, .cls files in `src/` directories
- Keep data files in CSV format in `data/` directories
- Build .xlam from source when releasing new versions

## Next Steps

After successful assembly:
1. Test all spelling and grammar features
2. Test all QS validation features
3. Test with sample BOQ files
4. Fix any bugs found
5. Update version number
6. Create release package
7. Distribute to users

## Support

If you encounter issues during assembly, refer to:
- `docs/developer-guides/Developer_Guide.md`
- `docs/developer-guides/Troubleshooting.md`
- Excel VBA documentation

---

**Author**: Excel Add-in Development Team
**Version**: 1.0.0
**Last Updated**: 2025-12-12
