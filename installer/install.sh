#!/bin/bash
# ==============================================================================
# Grammar & QS Checker Add-in Installer (macOS)
# ==============================================================================

echo ""
echo "========================================"
echo " Grammar & QS Checker Add-in Installer"
echo "========================================"
echo ""

# Define variables
ADDIN_NAME="GrammarChecker_QS.xlam"
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ADDIN_PATH="$SCRIPT_DIR/$ADDIN_NAME"
INSTALL_DIR="$HOME/Library/Application Support/Microsoft/Office/Excel/AddIns"

echo "Add-in file: $ADDIN_NAME"
echo "Install location: $INSTALL_DIR"
echo ""

# Check if add-in file exists
if [ ! -f "$ADDIN_PATH" ]; then
    echo "ERROR: Add-in file not found: $ADDIN_PATH"
    echo "Please ensure $ADDIN_NAME is in the same folder as this installer."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Create installation directory if it doesn't exist
if [ ! -d "$INSTALL_DIR" ]; then
    echo "Creating installation directory..."
    mkdir -p "$INSTALL_DIR"
fi

# Check if add-in already installed
if [ -f "$INSTALL_DIR/$ADDIN_NAME" ]; then
    echo ""
    echo "WARNING: An existing version of the add-in was found."
    read -p "Do you want to replace it? (y/n): " -n 1 -r
    echo ""
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo "Installation cancelled."
        read -p "Press Enter to exit..."
        exit 0
    fi
    echo "Removing old version..."
    rm "$INSTALL_DIR/$ADDIN_NAME"
fi

# Copy add-in to installation directory
echo ""
echo "Installing add-in..."
cp "$ADDIN_PATH" "$INSTALL_DIR/"

if [ $? -ne 0 ]; then
    echo "ERROR: Failed to copy add-in file."
    echo "Please check permissions and try again."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Add-in installed successfully!"
echo ""
echo "========================================"
echo " Installation Complete"
echo "========================================"
echo ""
echo "The Grammar & QS Checker add-in has been installed to:"
echo "$INSTALL_DIR/$ADDIN_NAME"
echo ""
echo "NEXT STEPS:"
echo "1. Restart Microsoft Excel if it's currently running"
echo "2. In Excel, go to: Tools > Excel Add-ins..."
echo "3. Check the box next to 'GrammarChecker_QS'"
echo "4. Click OK"
echo ""
echo "The add-in buttons will appear in the Excel ribbon."
echo ""
echo "If you encounter security warnings:"
echo "- Go to Excel > Preferences > Security & Privacy"
echo "- Enable macros from trusted sources"
echo ""
echo "For help, see the User Guide documentation."
echo ""

read -p "Press Enter to exit..."
