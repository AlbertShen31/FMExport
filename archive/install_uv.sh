#!/bin/bash
# Script to install uv on macOS

echo "Installing uv..."

# Method 1: Try official installer (recommended)
if curl -LsSf https://astral.sh/uv/install.sh | sh; then
    echo "✓ uv installed successfully via official installer"
    echo "Please restart your terminal or run: source ~/.cargo/env"
    exit 0
fi

# Method 2: Try with pip (requires --break-system-packages on macOS)
echo "Trying alternative installation method..."
if python3 -m pip install --break-system-packages uv; then
    echo "✓ uv installed successfully via pip"
    echo "Note: uv should be available in: ~/.local/bin/uv"
    echo "You may need to add ~/.local/bin to your PATH"
    exit 0
fi

# Method 3: Try Homebrew (if available and Xcode license accepted)
if command -v brew &> /dev/null; then
    echo "Trying Homebrew installation..."
    if brew install uv; then
        echo "✓ uv installed successfully via Homebrew"
        exit 0
    fi
fi

echo "✗ Installation failed. Please install manually:"
echo "  1. Visit: https://github.com/astral-sh/uv"
echo "  2. Or run: python3 -m pip install --break-system-packages uv"
echo "  3. Or run: brew install uv (if Homebrew is configured)"

