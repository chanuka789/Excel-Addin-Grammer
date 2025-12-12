# Step-by-Step Build Instructions for Excel Add-in

This guide will walk you through creating the **GrammarChecker_QS.xlam** Excel Add-in from the source files in this repository.

**Time Required**: 30-45 minutes
**Skill Level**: Intermediate Excel/VBA knowledge helpful but not required

---

## Prerequisites

- Microsoft Excel 2016 or later (Windows or Mac)
- The files in this repository (already cloned/downloaded)
- Basic familiarity with Excel

---

## Part 1: Create New Excel Workbook and Open VBA Editor

### Step 1: Create New Workbook

1. **Open Microsoft Excel**
2. **Create a new blank workbook** (File > New > Blank Workbook)
3. **Save immediately** with a temporary name:
   - Click **File > Save As**
   - Choose a location (e.g., your Desktop)
   - File name: `GrammarChecker_QS_Temp.xlsx`
   - Save as type: **Excel Workbook (*.xlsx)**
   - Click **Save**

### Step 2: Open VBA Editor

1. **Open the VBA Editor**:
   - **Windows**: Press `Alt + F11`
   - **Mac**: Press `Option + F11` or `Fn + Option + F11`

2. **You should see the VBA Editor window with**:
   - Left side: Project Explorer showing "VBAProject (GrammarChecker_QS_Temp.xlsx)"
   - Right side: Code window (may be empty)

3. **Set Project Name**:
   - In the VBA Editor, go to **Tools > VBAProject Properties**
   - In the "Project Name" field, type: `GrammarChecker_QS_AddIn`
   - Click **OK**

---

## Part 2: Import VBA Modules

### Step 3: Import Modules in Correct Order

**IMPORTANT**: Import modules in this exact order to avoid dependency errors.

1. **Right-click** on "VBAProject (GrammarChecker_QS_Temp.xlsx)" in the Project Explorer
2. Select **Import File...**
3. Navigate to the `src/modules/` folder in your downloaded repository

**Import these files in ORDER:**

#### Import Order (Critical - Follow Exactly):

**First** (Core Dependencies):
1. `ModUtility.bas` ← **MUST BE FIRST** (contains enums used by others)

**Second** (Infrastructure):
2. `ModLogging.bas`
3. `ModConfig.bas`

**Third** (Core Features):
4. `ModSpelling.bas`
5. `ModGrammar.bas`

**Fourth** (Main Module):
6. `ModMain.bas`

**Fifth** (QS Modules - any order):
7. `ModQSValidator.bas`
8. `ModQSDictionary.bas`
9. `ModBOQAnalysis.bas`
10. `ModUnitValidator.bas`
11. `ModCostAnalysis.bas`
12. `ModDescriptionAnalysis.bas`
13. `ModFIDIC.bas`

**How to Import Each File:**
1. Right-click **VBAProject** > **Import File...**
2. Navigate to `src/modules/`
3. Select the .bas file (e.g., `ModUtility.bas`)
4. Click **Open**
5. You should see the module appear in the Project Explorer under "Modules"
6. Repeat for each file in order

**After importing, your Project Explorer should show:**
```
VBAProject (GrammarChecker_QS_Temp.xlsx)
├── Microsoft Excel Objects
│   └── ThisWorkbook
└── Modules
    ├── ModUtility
    ├── ModLogging
    ├── ModConfig
    ├── ModSpelling
    ├── ModGrammar
    ├── ModMain
    ├── ModQSValidator
    ├── ModQSDictionary
    ├── ModBOQAnalysis
    ├── ModUnitValidator
    ├── ModCostAnalysis
    ├── ModDescriptionAnalysis
    └── ModFIDIC
```

### Step 4: Verify Modules Imported Correctly

1. In the VBA Editor, go to **Debug > Compile VBAProject**
2. If there are no errors, you're good! ✅
3. If you see errors like "User-defined type not defined":
   - You imported modules in wrong order
   - Delete all modules and re-import in correct order

---

## Part 3: Create Hidden Data Worksheets

### Step 5: Add Worksheets for Data Storage

**Return to Excel** (close VBA Editor or press `Alt + F11` again)

You should see your workbook with the default sheets (Sheet1, Sheet2, Sheet3).

**Create the following worksheets:**

1. **Click the "+" button** (or right-click sheet tab > Insert > Worksheet) to add a new sheet
2. **Rename it** by double-clicking the sheet tab

**Create these 10 worksheets** (exact names are critical):

1. `Dictionary_EN`
2. `GrammarRules`
3. `Settings`
4. `QS_Dictionary`
5. `QS_UnitMasters`
6. `QS_FIDICReferences`
7. `QS_ItemTemplates`
8. `QS_DescriptionPatterns`
9. `QS_Settings`
10. `ChangeLog`

**How to create each:**
1. Click the **+** or **New Sheet** button (at bottom left of Excel)
2. **Double-click** the new sheet tab (e.g., "Sheet4")
3. Type the exact name (e.g., `Dictionary_EN`)
4. Press **Enter**
5. Repeat for all 10 worksheets

**After creating all, you should have:**
- Dictionary_EN
- GrammarRules
- Settings
- QS_Dictionary
- QS_UnitMasters
- QS_FIDICReferences
- QS_ItemTemplates
- QS_DescriptionPatterns
- QS_Settings
- ChangeLog
- (Plus any default sheets like Sheet1, Sheet2, etc. - you can delete these later)

---

## Part 4: Import CSV Data into Worksheets

### Step 6: Import Dictionary Data

**For each worksheet, you'll import the corresponding CSV file.**

#### 6.1 Import Dictionary_EN

1. **Click on the `Dictionary_EN` sheet tab**
2. **Click cell A1**
3. Go to **Data > Get Data > From File > From Text/CSV** (Excel 2016+)
   - **OR** go to **Data > From Text/CSV** (newer versions)
   - **OR** use **File > Import** on Mac
4. Navigate to `data/dictionaries/english_dictionary.csv`
5. Click **Import** or **Open**
6. In the preview dialog:
   - Ensure "Delimiter" is set to **Comma**
   - Click **Load** (NOT "Load To...")
7. The data should appear in the worksheet starting at A1

**Verify you see headers:**
- A1: Word
- B1: Length
- C1: Frequency
- D1: Category

#### 6.2 Import GrammarRules

1. **Click on the `GrammarRules` sheet tab**
2. **Click cell A1**
3. Go to **Data > From Text/CSV**
4. Navigate to `data/grammar-rules/grammar_rules.csv`
5. Click **Import/Open**
6. Click **Load**

**Verify headers:**
- A1: RuleID
- B1: Pattern
- C1: Replacement
- D1: Severity
- E1: Category
- F1: Description

#### 6.3 Import QS_Dictionary

1. **Click on the `QS_Dictionary` sheet tab**
2. **Click cell A1**
3. Go to **Data > From Text/CSV**
4. Navigate to `data/qs-data/qs_terminology.csv`
5. Click **Import/Open**
6. Click **Load**

**Verify headers:**
- A1: Term
- B1: CorrectSpelling
- C1: CommonMisspellings
- D1: Category
- E1: Definition
- F1: StandardUnit
- G1: RegionalVariants
- H1: RelatedTerms

#### 6.4 Import QS_UnitMasters

1. **Click on the `QS_UnitMasters` sheet tab**
2. **Click cell A1**
3. Go to **Data > From Text/CSV**
4. Navigate to `data/qs-data/unit_masters.csv`
5. Click **Import/Open**
6. Click **Load**

**Verify headers:**
- A1: UnitCode
- B1: UnitName
- C1: ApplicableItems
- D1: ConversionFactor
- E1: Precision
- F1: CommonMisspellings

#### 6.5 Import QS_FIDICReferences

1. **Click on the `QS_FIDICReferences` sheet tab**
2. **Click cell A1**
3. Go to **Data > From Text/CSV**
4. Navigate to `data/qs-data/fidic_references.csv`
5. Click **Import/Open**
6. Click **Load**

**Verify headers:**
- A1: ClauseNumber
- B1: ClauseTitle
- C1: Requirements
- D1: RelatedClauses
- E1: PaymentRelated
- F1: TimelineRelated

#### 6.6 Setup Settings Worksheets (Manual Headers)

For `Settings`, `QS_Settings`, and `ChangeLog`, we'll just add headers manually:

**Settings Sheet:**
1. Click on the `Settings` tab
2. In cell A1, type: `SettingName`
3. In cell B1, type: `SettingValue`
4. In cell C1, type: `SettingType`
5. In cell D1, type: `Description`

**QS_Settings Sheet:**
1. Click on the `QS_Settings` tab
2. Add same headers as Settings:
   - A1: `SettingName`
   - B1: `SettingValue`
   - C1: `SettingType`
   - D1: `Description`

**ChangeLog Sheet:**
1. Click on the `ChangeLog` tab
2. Add these headers:
   - A1: `Timestamp`
   - B1: `Level`
   - C1: `Location`
   - D1: `ErrorType`
   - E1: `Original`
   - F1: `Corrected`
   - G1: `Severity`

**QS_ItemTemplates and QS_DescriptionPatterns:**
For now, just add headers:

**QS_ItemTemplates:**
- A1: `ItemType`
- B1: `TemplateFormat`
- C1: `RequiredElements`
- D1: `Example`

**QS_DescriptionPatterns:**
- A1: `PatternID`
- B1: `ItemCategory`
- C1: `RegexPattern`
- D1: `RequiredFields`
- E1: `ValidationLevel`

---

## Part 5: Hide Data Worksheets

### Step 7: Hide All Data Worksheets

Now we'll hide these worksheets so users don't see them.

**Option A: Hide Normally (Easy to unhide)**

1. **Right-click** on the `Dictionary_EN` sheet tab
2. Select **Hide**
3. Repeat for ALL data sheets:
   - Dictionary_EN
   - GrammarRules
   - Settings
   - QS_Dictionary
   - QS_UnitMasters
   - QS_FIDICReferences
   - QS_ItemTemplates
   - QS_DescriptionPatterns
   - QS_Settings
   - ChangeLog

**Option B: Hide Deeply (Recommended - Can't unhide from Excel UI)**

1. **Open VBA Editor** (Alt + F11)
2. In the **Immediate Window** (View > Immediate Window or Ctrl+G), type these commands **one at a time** and press Enter after each:

```vba
ThisWorkbook.Worksheets("Dictionary_EN").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("GrammarRules").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("Settings").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_Dictionary").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_UnitMasters").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_FIDICReferences").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_ItemTemplates").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_DescriptionPatterns").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("QS_Settings").Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets("ChangeLog").Visible = xlSheetVeryHidden
```

**Note:** If you get an error "Subscript out of range", double-check the worksheet name spelling.

3. **Return to Excel** (Alt + F11)
4. **Delete any remaining default sheets** (Sheet1, Sheet2, etc.) - right-click > Delete
5. You should now have **NO visible worksheets** - this is OK for an add-in!

---

## Part 6: Configure ThisWorkbook Module

### Step 8: Add Initialization Code to ThisWorkbook

1. **Open VBA Editor** (Alt + F11)
2. In the **Project Explorer**, double-click **ThisWorkbook** under "Microsoft Excel Objects"
3. In the code window that appears, **paste this code:**

```vba
Option Explicit

'==============================================================================
' ThisWorkbook - Add-in Initialization
'==============================================================================

Private Sub Workbook_Open()
    ' This runs when the add-in loads with Excel
    On Error Resume Next
    Call InitializeAddIn
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    ' This runs when add-in is first installed
    On Error Resume Next
    Call OnInstall
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinUninstall()
    ' This runs when add-in is uninstalled
    On Error Resume Next
    Call OnUninstall
    On Error GoTo 0
End Sub

Private Sub InitializeAddIn()
    ' Initialize all modules

    ' Load settings
    Call ModConfig.LoadSettings

    ' Initialize logging
    Call ModLogging.InitializeLogging

    ' Initialize spelling dictionary
    Call ModSpelling.InitializeDictionary

    ' Initialize grammar rules
    Call ModGrammar.InitializeGrammarRules

    ' Initialize QS modules if enabled
    If ModConfig.EnableQSValidation Then
        Call ModQSValidator.InitializeQS
    End If

    ' Log successful startup
    Call ModLogging.LogEvent("Add-in initialized successfully v" & ModMain.ADDIN_VERSION, "INFO")
End Sub

Private Sub OnInstall()
    ' First-time installation
    Call ModConfig.CreateDefaultSettings

    MsgBox "Grammar & QS Checker Add-in v" & ModMain.ADDIN_VERSION & " installed successfully!" & vbCrLf & vbCrLf & _
           "Look for the 'Grammar & QS' tab in the Excel ribbon." & vbCrLf & vbCrLf & _
           "For help, click the Help button on the ribbon.", _
           vbInformation, "Installation Complete"
End Sub

Private Sub OnUninstall()
    MsgBox "Grammar & QS Checker Add-in has been uninstalled.", vbInformation, "Uninstall Complete"
End Sub
```

4. **Save** (Ctrl + S)

---

## Part 7: Save as Excel Add-in (.xlam)

### Step 9: Save as .xlam Format

1. **In VBA Editor**, go to **Debug > Compile VBAProject**
   - This checks for any errors
   - If you see errors, fix them before continuing

2. **Close VBA Editor** (or press Alt + F11 to return to Excel)

3. **In Excel**, go to **File > Save As**

4. **Choose save location**:
   - For testing: Save to your Desktop or Documents folder
   - For installation: You can move it later

5. **Important Settings:**
   - **File name**: `GrammarChecker_QS.xlam`
   - **Save as type**: Select **"Excel Add-in (*.xlam)"** from dropdown
     - **Windows**: Look for "Excel Add-in (*.xlam)"
     - **Mac**: Look for "Excel Add-In (.xlam)"

6. **Click Save**

7. **Excel will close the file automatically** - this is normal for add-ins!

**You now have created: `GrammarChecker_QS.xlam`** ✅

---

## Part 8: Install and Test the Add-in

### Step 10: Install the Add-in in Excel

**Method 1: Manual Installation (Recommended for Testing)**

1. **Open Excel** (new blank workbook)

2. **Go to Add-ins Manager**:
   - **Windows**: File > Options > Add-ins
   - **Mac**: Tools > Excel Add-ins...

3. **On Windows**:
   - At the bottom, select **"Excel Add-ins"** from the "Manage" dropdown
   - Click **Go...**

4. **In the Add-ins dialog**:
   - Click **Browse...**
   - Navigate to where you saved `GrammarChecker_QS.xlam`
   - Select the file and click **OK**

5. **Enable the Add-in**:
   - You should see "GrammarChecker_QS" in the list
   - **Check the box** next to it
   - Click **OK**

6. **Check for Security Warning**:
   - If you see a security warning about macros:
   - Click **Enable Macros** or **Enable Content**

**Method 2: Copy to AddIns Folder (Permanent Installation)**

1. **Copy** `GrammarChecker_QS.xlam` to:
   - **Windows**: `C:\Users\[YourUsername]\AppData\Roaming\Microsoft\AddIns\`
   - **Mac**: `~/Library/Application Support/Microsoft/Office/Excel/AddIns/`

2. **Follow steps 2-6 from Method 1 above**

---

### Step 11: Verify Installation

**Check for Ribbon Tab:**

1. **Look at the Excel Ribbon** (top of Excel window)
2. You should see a new tab called **"Grammar & QS"**
3. If you DON'T see it:
   - The add-in might not be enabled (go back to Step 10)
   - There might be a macro security issue (see Troubleshooting below)

**If you see "Grammar & QS" tab:** ✅ Success!

---

## Part 9: Test the Add-in

### Step 12: Test Basic Functionality

**Test 1: Spelling Check**

1. **Create a new Excel workbook**
2. **Type some text with spelling errors**:
   ```
   A1: This is a test
   A2: This has a speling error
   A3: Another tets here
   ```

3. **Select cells A1:A3**

4. **Click the "Grammar & QS" tab**

5. **Click "Check Now" button**

6. **You should see**:
   - A message box saying "Found X error(s)"
   - List of errors detected
   - "speling" → "spelling"
   - "tets" → "test"

**If you see errors detected:** ✅ Spelling check works!

**Test 2: Grammar Check**

1. **Type text with grammar errors**:
   ```
   A1: This  has  double  spaces
   A2: No space after comma,like this
   A3: Space before period .
   ```

2. **Select cells A1:A3**

3. **Click "Check Now"**

4. **You should see**:
   - Multiple spaces detected
   - Comma spacing issues
   - Period spacing issues

**If you see grammar errors detected:** ✅ Grammar check works!

**Test 3: QS Validation (Basic)**

1. **Create a simple BOQ**:
   ```
   A1: Description          B1: Unit    C1: Quantity    D1: Rate    E1: Amount
   A2: Concrete Work        B2: M³      C2: 10          D2: 100     E2: 1000
   A3: Steel Reinforcement  B3:         C3: 5           D3: 200     D3: 1000
   A4: Excavation          B4: M³      C4: 20          D4: 50      E4: 999
   ```

2. **Select range A1:E4**

3. **Click "Grammar & QS" tab > "Validate BOQ" button**

4. **You should see**:
   - Missing unit in B3 detected ✅
   - Calculation error in E4 (should be 1000, not 999) ✅

**If you see these errors:** ✅ QS validation works!

---

## Troubleshooting

### Problem: "Grammar & QS" tab doesn't appear

**Solution 1: Enable Add-in**
- File > Options > Add-ins > Manage Excel Add-ins > Go
- Check the box next to GrammarChecker_QS
- Click OK

**Solution 2: Enable Macros**
- File > Options > Trust Center > Trust Center Settings
- Macro Settings > Enable all macros (for testing)
- Click OK
- Restart Excel

**Solution 3: Check Installation**
- File > Options > Add-ins
- Look in "Active Application Add-ins" section
- GrammarChecker_QS should be listed as "Active"

### Problem: "Compile Error" when opening

**Solution:**
- You imported modules in wrong order
- Rebuild: Delete all modules, re-import in correct order (ModUtility FIRST)

### Problem: "Subscript out of range" error

**Solution:**
- Worksheet names don't match exactly
- Check worksheet names are EXACT (case-sensitive)
- Example: "Dictionary_EN" not "Dictionary_en"

### Problem: "Can't find dictionary" warning

**Solution:**
- CSV data wasn't imported correctly
- Check Dictionary_EN worksheet has data
- Ensure column A has words starting from A2 (A1 is header "Word")

### Problem: Very slow performance

**Solution:**
- Dictionary might not be loading into memory
- Check Immediate Window for errors: Ctrl+G in VBA Editor
- Add Debug.Print statements to InitializeDictionary

### Problem: No results showing

**Solution:**
- Results currently show in MsgBox (basic)
- If nothing shows: Check g_ErrorCollection count in Immediate Window
- Type: `? ModMain.g_ErrorCollection.Count` after running check

---

## Advanced: Testing with Debug Mode

### Enable Debug Output

1. **Open VBA Editor** (Alt + F11)
2. **Open Immediate Window** (Ctrl + G or View > Immediate Window)
3. **Run a check in Excel**
4. **Switch back to VBA Editor**
5. **Look in Immediate Window** for any debug messages

### Test Individual Functions

In Immediate Window, you can test functions directly:

```vba
' Test spelling check
? ModSpelling.IsWordSpelledCorrectly("test")
' Should return: True

? ModSpelling.IsWordSpelledCorrectly("tset")
' Should return: False

' Test Levenshtein distance
? ModUtility.LevenshteinDistance("kitten", "sitting")
' Should return: 3

' Test dictionary loaded
? ModSpelling.GetDictionaryWordCount()
' Should return: 150+ (number of words in dictionary)

' Test QS term
? ModQSDictionary.IsQSTerm("concrete")
' Should return: True
```

---

## Next Steps After Successful Testing

### 1. Expand Dictionaries
- Add more words to `data/dictionaries/english_dictionary.csv`
- Expand to 3,000-5,000 words
- Re-import into Dictionary_EN worksheet

### 2. Add More Grammar Rules
- Edit `data/grammar-rules/grammar_rules.csv`
- Add custom rules for your needs
- Re-import into GrammarRules worksheet

### 3. Customize QS Data
- Add company-specific terms to QS_Dictionary
- Add custom units to QS_UnitMasters
- Update FIDIC references if using different version

### 4. Distribute to Users
- Copy .xlam file to shared network location
- Create distribution package with installer scripts
- Provide Quick Start Guide to users

### 5. Future Enhancements
- Add UserForms for better UI (frmResults, frmSettings)
- Implement real-time checking
- Add export to PDF functionality
- Create custom reports

---

## Summary Checklist

✅ Created new Excel workbook
✅ Opened VBA Editor
✅ Imported 13 VBA modules in correct order
✅ Created 10 hidden data worksheets
✅ Imported CSV data into worksheets
✅ Added headers to settings worksheets
✅ Hidden all data worksheets (xlSheetVeryHidden)
✅ Configured ThisWorkbook with initialization code
✅ Compiled VBA project (no errors)
✅ Saved as .xlam format
✅ Installed add-in in Excel
✅ Verified "Grammar & QS" tab appears
✅ Tested spelling check
✅ Tested grammar check
✅ Tested QS validation

**Congratulations! You've successfully built the Excel Add-in!** 🎉

---

## Support

If you encounter issues:
1. Check the Troubleshooting section above
2. Review ASSEMBLY_GUIDE.md for additional details
3. Check docs/developer-guides/Developer_Guide.md for technical info
4. Verify all CSV files imported correctly

For further assistance, refer to the project documentation or create an issue in the repository.

---

**Version**: 1.0.0
**Last Updated**: 2025-12-12
