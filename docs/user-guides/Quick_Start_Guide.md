# Quick Start Guide - Grammar & QS Checker Add-in

## Welcome!

Thank you for installing the Grammar & QS Checker Add-in for Microsoft Excel. This guide will help you get started quickly.

## What This Add-in Does

The Grammar & QS Checker provides two main sets of features:

### 1. **Spelling & Grammar Checking**
- Detects spelling errors in your spreadsheets
- Identifies common grammar mistakes
- Suggests corrections
- Works with any Excel workbook

### 2. **QS (Quantity Surveying) / BOQ Validation**
- Validates Bill of Quantities (BOQ) structure
- Checks for missing data (quantities, rates, units)
- Validates calculations (Qty × Rate = Amount)
- Detects cost anomalies and unusual rates
- Validates construction terminology
- Checks unit consistency
- FIDIC 1999 clause reference validation

---

## First Time Setup

After installing the add-in:

1. **Open Microsoft Excel**
2. Look for a new tab in the ribbon called **"Grammar & QS"**
3. If you don't see it:
   - Go to `File` > `Options` > `Add-ins`
   - At the bottom, select "Excel Add-ins" and click "Go..."
   - Check the box next to "GrammarChecker_QS"
   - Click OK

---

## Basic Usage

### Checking Spelling & Grammar

1. **Open your Excel workbook**
2. **Select the cells** you want to check (or don't select anything to check the entire sheet)
3. **Click the "Check Now" button** on the Grammar & QS tab
4. Review the results dialog showing any errors found
5. **Accept or reject** each suggestion

**Quick Options:**
- Use the **"Check Options"** dropdown to check spelling only, grammar only, or both

### Validating a BOQ (Bill of Quantities)

1. **Open your BOQ workbook**
2. **Select the BOQ range** (including headers)
3. **Click "Validate BOQ"** button on the Grammar & QS tab
4. The add-in will check for:
   - Missing quantities, rates, or units
   - Incorrect calculations
   - Invalid units
   - Cost anomalies
   - Description issues
5. Review the comprehensive report
6. **Apply corrections** as needed

**QS Quick Checks:**
- **Quick Check**: Fast scan for missing critical data
- **Cost Analysis**: Identify unusual rates and outliers
- **Unit Validation**: Check unit consistency
- **Description Check**: Find duplicates and incomplete descriptions

---

## Understanding Results

Results are categorized by severity:

- **🔴 Critical**: Must be fixed (e.g., missing data, calculation errors)
- **🟡 Warning**: Should be reviewed (e.g., unusual rates, potential typos)
- **🔵 Info**: Informational (e.g., suggestions for improvement)

---

## Tips for Best Results

1. **Select the Right Range**: For BOQ validation, include the header row with column names (Description, Unit, Quantity, Rate, Amount)

2. **Review Before Accepting**: Not all suggestions may be correct for your context - review each one

3. **Use Settings**: Configure thresholds and options via the Settings button to match your needs

4. **Regular Checks**: Run validation periodically during BOQ development to catch errors early

5. **Save Your Work**: Always save your workbook before accepting mass corrections

---

## Common Tasks

### Add a Word to Dictionary
If a word is flagged as misspelled but is correct (e.g., company names, technical terms):
1. Right-click on the error in results
2. Select "Add to Dictionary"
3. The word won't be flagged again

### Adjust Cost Anomaly Threshold
If you're getting too many cost anomaly warnings:
1. Click "QS Settings"
2. Adjust "Cost Anomaly Threshold %" (default: 50%)
3. Higher values = fewer warnings

### Check Multiple Worksheets
1. Select the first worksheet tab
2. Hold Ctrl (Windows) or Cmd (Mac)
3. Click additional worksheet tabs
4. Click "Check Now"
5. All selected sheets will be checked

---

## Keyboard Shortcuts

_(These may be configured in future versions)_

- `Ctrl+Shift+G`: Check Grammar (planned)
- `Ctrl+Shift+S`: Check Spelling (planned)
- `Ctrl+Shift+Q`: Validate BOQ (planned)

---

## Troubleshooting

**Problem**: Add-in doesn't appear in ribbon
- **Solution**: Go to File > Options > Add-ins, ensure add-in is enabled

**Problem**: "Macro security warning" appears
- **Solution**: Go to File > Options > Trust Center > Macro Settings, enable macros

**Problem**: Very slow on large spreadsheets
- **Solution**: Select smaller ranges rather than entire worksheet

**Problem**: Too many false positives
- **Solution**: Adjust settings, add custom words to dictionary

**Problem**: BOQ validation doesn't work
- **Solution**: Ensure your BOQ has proper column headers (Description, Unit, Quantity, Rate, Amount)

---

## Getting Help

- **Help Button**: Click the "Help" button on the ribbon for quick tips
- **User Guides**: See detailed guides in `docs/user-guides/`
- **Support**: Report issues or request features in the project repository

---

## Next Steps

- Explore the **Settings** dialog to customize behavior
- Read the **QS Features Guide** for advanced BOQ validation
- Check out **Developer Guide** if you want to extend the add-in

---

**Version**: 1.0.0
**Last Updated**: 2025-12-12

---

Enjoy using the Grammar & QS Checker Add-in!
