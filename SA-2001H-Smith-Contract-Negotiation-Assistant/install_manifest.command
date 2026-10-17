#!/bin/bash
# Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.
#
# macOS sideload installer (for LOCAL TESTING and SINGLE-MACHINE installs).
#
# This script (1) stamps your server URL from .env into the manifest, then
# (2) copies the manifest into Microsoft Word's local add-in folder.
#
# For organization-wide rollout, use centralized deployment via the Microsoft 365
# admin center instead (see the user guide) — that path is OS-agnostic and does not
# use this script.

echo "====================================================="
echo "  Contract Negotiation Assistant - macOS Sideload    "
echo "====================================================="

# Resolve folder locations relative to this script (Contract-Frontend/).
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
BACKEND_DIR="$( cd "$DIR/../Contract-Backend" && pwd )"

# --- Step 1: Stamp the server URL from .env into manifest.xml (and taskpane.js) ---
if [ -f "$BACKEND_DIR/setup.js" ]; then
    echo "Stamping server URL from .env into the manifest..."
    ( cd "$BACKEND_DIR" && node setup.js )
    if [ $? -ne 0 ]; then
        echo "CRITICAL ERROR: Manifest stamping failed. Fix the error above (usually NODE_SERVER_URL in .env), then re-run."
        echo "Press any key to close this window..."
        read -n 1
        exit 1
    fi
else
    echo "WARNING: setup.js not found in Contract-Backend; attempting to use an existing manifest.xml as-is."
fi

# --- Step 2: Copy the generated manifest into Word's sandbox add-in folder ---
TARGET_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
if [ ! -d "$TARGET_DIR" ]; then
    echo "Creating Microsoft Word Sandbox Add-in directory..."
    mkdir -p "$TARGET_DIR"
fi

if [ -f "$DIR/manifest.xml" ]; then
    cp "$DIR/manifest.xml" "$TARGET_DIR/manifest.xml"
    echo "SUCCESS: Add-In Manifest cleanly sideloaded onto this workstation."
    echo "Please completely restart Microsoft Word to activate the tool ribbon controls."
else
    echo "CRITICAL ERROR: manifest.xml was not found after stamping."
    echo "Expected path: $DIR/manifest.xml"
fi

echo "Press any key to close this window..."
read -n 1
exit
