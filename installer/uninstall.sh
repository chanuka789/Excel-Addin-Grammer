#!/bin/bash
# ==============================================================================
# Grammar & QS Checker Add-in Uninstaller (macOS)
# ==============================================================================

echo ""
echo "=========================================="
echo " Grammar & QS Checker Add-in Uninstaller"
echo "=========================================="
echo ""

ADDIN_NAME="GrammarChecker_QS.xlam"
INSTALL_DIR="$HOME/Library/Application Support/Microsoft/Office/Excel/AddIns"

echo "This will remove the Grammar & QS Checker add-in from:"
echo "$INSTALL_DIR/$ADDIN_NAME"
echo ""

read -p "Do you want to continue? (y/n): " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Uninstallation cancelled."
    read -p "Press Enter to exit..."
    exit 0
fi

echo ""
echo "Removing add-in..."

if [ -f "$INSTALL_DIR/$ADDIN_NAME" ]; then
    rm "$INSTALL_DIR/$ADDIN_NAME"
    echo "Add-in removed successfully!"
else
    echo "Add-in file not found. It may have been already removed."
fi

echo ""
echo "=========================================="
echo " Uninstallation Complete"
echo "=========================================="
echo ""
echo "The add-in has been removed from your system."
echo ""
echo "If Excel is running, please restart it."
echo ""
echo "To remove the add-in from Excel's list:"
echo "1. Open Excel"
echo "2. Go to: Tools > Excel Add-ins..."
echo "3. Uncheck 'GrammarChecker_QS' if it appears"
echo "4. Click OK"
echo ""

read -p "Press Enter to exit..."
