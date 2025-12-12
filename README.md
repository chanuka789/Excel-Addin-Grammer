# Excel Spelling & Grammar + QS Validation Add-in

A professional Excel Add-in (.xlam) that provides automatic spelling and grammar correction, plus comprehensive Quantity Surveying (QS) features for BOQ validation and analysis.

## Features

### Core Features
- **Spelling Check**: Detects and corrects spelling errors with 3,000+ word dictionary
- **Grammar Check**: Identifies common grammar mistakes (punctuation, spacing, capitalization)
- **Custom Ribbon**: Integrated toolbar buttons in Excel ribbon
- **Multi-Workbook Support**: Works with any open Excel file
- **Change Tracking**: Complete audit trail with undo functionality

### QS (Quantity Surveying) Features
- **BOQ Analysis**: Structure validation and completeness checking
- **QS Terminology**: 2,000+ construction term dictionary with auto-correction
- **Unit Validation**: Standardization and consistency checking
- **Cost Analysis**: Rate validation and outlier detection
- **Description Standardization**: BOQ description formatting and duplicate detection
- **FIDIC Validation**: FIDIC 1999 clause reference checking
- **IPC Support**: Interim Payment Certificate template validation
- **Calculation Verification**: Quantity × Rate = Amount validation
- **Summary Cross-Checking**: Verify summary totals against detail rows
- **Comprehensive Reporting**: QS validation reports with PDF export

## Project Structure

```
Excel-Addin-Grammer/
├── src/
│   ├── modules/          # VBA modules (.bas files)
│   ├── forms/            # UserForm files (.frm)
│   └── classes/          # Class modules (.cls)
├── data/
│   ├── dictionaries/     # Spelling dictionary data
│   ├── qs-data/          # QS terminology, units, FIDIC references
│   └── grammar-rules/    # Grammar validation rules
├── docs/
│   ├── user-guides/      # End-user documentation
│   └── developer-guides/ # Technical documentation
├── installer/            # Installation scripts
├── ribbon/               # Custom ribbon XML configuration
├── tests/                # Test files and test data
└── ASSEMBLY_GUIDE.md     # Instructions to build the .xlam file
```

## Quick Start

### For Developers
1. Read `ASSEMBLY_GUIDE.md` for step-by-step instructions to build the Add-in
2. Import VBA modules from `src/modules/` into Excel
3. Import data from `data/` into hidden worksheets
4. Configure ribbon using `ribbon/customUI.xml`
5. Save as .xlam (Excel Add-in format)

### For End Users
1. Download the pre-built `GrammarChecker_QS.xlam` file
2. Run `installer/install.bat` (Windows) or `installer/install.sh` (Mac)
3. Restart Excel
4. Look for "Spelling & Grammar" and "BOQ & QS Analysis" buttons in the ribbon

## Installation Paths

- **Windows**: `C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\`
- **Mac**: `~/Library/Application Support/Microsoft/Office/Excel/AddIns/`

## System Requirements

- Microsoft Excel 2016 or later
- Windows 10/11 or macOS 10.13+
- Macro security set to allow digitally signed macros

## Documentation

- **User Guide (Grammar)**: `docs/user-guides/Grammar_User_Guide.md`
- **User Guide (QS Features)**: `docs/user-guides/QS_User_Guide.md`
- **Developer Guide**: `docs/developer-guides/Developer_Guide.md`
- **Assembly Guide**: `ASSEMBLY_GUIDE.md`

## License

Free to use for personal and commercial purposes.

## Version

**Current Version**: 1.0.0 (Initial Release)

## Support

For issues, questions, or feature requests, please open an issue in this repository.
