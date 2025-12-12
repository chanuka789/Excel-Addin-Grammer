# Changelog

All notable changes to the Grammar & QS Checker Add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-12

### Added - Initial Release

#### Core Features
- **Spelling Check Engine**
  - Dictionary-based spell checking with 150+ base words (expandable to 3,000-5,000)
  - Levenshtein distance algorithm for smart suggestions
  - Custom dictionary support (add your own words)
  - Case-insensitive matching
  - Configurable settings

- **Grammar Check Engine**
  - Rule-based grammar validation
  - Built-in rules for common errors (double spaces, punctuation, etc.)
  - Extensible rule system via CSV import
  - Severity levels (Critical, Warning, Info)
  - Category-based error classification

- **Change Tracking & Undo**
  - Complete audit trail of all corrections
  - Undo functionality (revert last N changes)
  - Change log export to CSV
  - Timestamps and user actions logged

#### QS (Quantity Surveying) Features
- **BOQ Analysis**
  - Structure validation (headers, format)
  - Missing data detection (quantities, rates, units, descriptions)
  - Calculation validation (Qty × Rate = Amount)
  - Duplicate description detection
  - Summary sheet cross-checking

- **QS Terminology Dictionary**
  - 200+ construction/QS terms
  - Common misspelling detection and correction
  - Regional variant support (UK, US, AU)
  - Standard unit associations
  - Category-based organization

- **Unit Validation**
  - Standard unit master list (M, M², M³, NO, TONNE, etc.)
  - Unit consistency checking
  - Common misspelling corrections
  - Applicable item type validation

- **Cost Analysis**
  - Zero and negative rate detection
  - Statistical outlier detection (configurable threshold)
  - Rate anomaly identification
  - Mean and standard deviation calculations

- **Description Analysis**
  - Incomplete description detection
  - Duplicate description identification
  - Length validation
  - Future: Pattern-based validation

- **FIDIC Validation**
  - Complete FIDIC 1999 clause reference database (200+ clauses)
  - Automatic clause number extraction from text
  - Clause validation against official database
  - Related clause suggestions

#### User Interface
- **Custom Ribbon Tab**
  - "Grammar & QS" tab in Excel ribbon
  - Spelling & Grammar group with Check Now button
  - QS Validation group with Validate BOQ button
  - Settings and Help buttons
  - Dropdown menus for quick options

- **Configuration**
  - Settings dialog for core features
  - QS-specific settings dialog
  - Persistent settings storage
  - Threshold customization (cost anomalies, etc.)

- **Progress Indicators**
  - Status bar messages during scanning
  - Cell count and progress tracking
  - Cancellable operations

#### Data Files
- English dictionary (CSV format, 150+ words)
- Grammar rules (CSV format, 17 rules)
- QS terminology (CSV format, 200+ terms)
- Unit masters (CSV format, 40+ units)
- FIDIC references (CSV format, 200+ clauses)

#### Installation & Deployment
- Windows installer script (.bat)
- macOS installer script (.sh)
- Windows uninstaller script
- macOS uninstaller script
- Automated AddIns folder detection
- One-click installation

#### Documentation
- Comprehensive README
- Assembly Guide (step-by-step .xlam creation)
- Quick Start Guide (user-facing)
- Developer Guide (technical documentation)
- Inline code documentation
- CSV data file documentation

#### VBA Modules
- ModUtility (helper functions, enums)
- ModLogging (change tracking, undo)
- ModConfig (settings management)
- ModSpelling (spelling check engine)
- ModGrammar (grammar check engine)
- ModMain (entry points, orchestration)
- ModQSValidator (QS orchestrator)
- ModQSDictionary (QS terminology)
- ModBOQAnalysis (BOQ validation)
- ModUnitValidator (unit validation)
- ModCostAnalysis (cost/rate analysis)
- ModDescriptionAnalysis (description checking)
- ModFIDIC (FIDIC clause validation)

### Technical Details
- Language: VBA (Visual Basic for Applications)
- Platform: Microsoft Excel 2016+
- Format: .xlam (Excel Add-in)
- Architecture: Modular, extensible design
- Performance: Optimized for large datasets (10,000+ cells)
- Security: Local processing, no external data transmission

### Known Limitations
- Dictionary size is limited (start with 150 words, expand manually)
- Grammar rules are pattern-based (not NLP-based)
- QS validation assumes standard BOQ format
- No real-time checking (manual trigger required)
- No multi-language support yet (English only)
- No UserForms implemented in v1.0 (results shown via MsgBox)

### Future Roadmap (Post-1.0)
- UserForms for results display (frmResults, frmSettings, etc.)
- Real-time checking with Application events
- Expanded dictionaries (3,000-5,000 words)
- Advanced QS reporting with charts
- BOQ template support (IPC, VO templates)
- Historical BOQ comparison
- Multi-workbook BOQ consolidation
- Machine learning for smarter suggestions
- Multi-language support
- Custom user-defined rules interface
- Export to PDF functionality

---

## [Unreleased]

### Planned for v1.1
- UserForm-based results dialog
- Enhanced reporting
- Extended dictionaries
- Performance improvements
- Bug fixes based on user feedback

---

## Version History Summary

- **v1.0.0** (2025-12-12): Initial release with core spelling/grammar and QS features

---

For detailed information, see:
- [README.md](README.md) - Project overview
- [ASSEMBLY_GUIDE.md](ASSEMBLY_GUIDE.md) - Build instructions
- [docs/user-guides/](docs/user-guides/) - User documentation
- [docs/developer-guides/](docs/developer-guides/) - Developer documentation
