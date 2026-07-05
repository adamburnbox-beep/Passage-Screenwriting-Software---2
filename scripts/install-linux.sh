#!/usr/bin/env bash
# Publish the Linux app and install a desktop launcher for the current user.
# Run this again after any code change to update the installed copy.
#
# Usage:
#   scripts/install-linux.sh                 # framework-dependent (needs dotnet runtime)
#   scripts/install-linux.sh --self-contained  # bundles the runtime (bigger, no dotnet needed)
#   scripts/install-linux.sh --uninstall
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROJECT="$REPO_DIR/Passage/Passage.App.Linux/Passage.App.Linux.csproj"
APP_DIR="$HOME/.local/share/passage"
BIN_DIR="$HOME/.local/bin"
DESKTOP_FILE="$HOME/.local/share/applications/passage.desktop"

if [[ "${1:-}" == "--uninstall" ]]; then
    rm -rf "$APP_DIR"
    rm -f "$BIN_DIR/passage" "$DESKTOP_FILE"
    command -v update-desktop-database >/dev/null && update-desktop-database "$HOME/.local/share/applications" || true
    echo "Passage uninstalled."
    exit 0
fi

echo "Publishing Passage (Release)..."
dotnet publish "$PROJECT" -c Release -o "$APP_DIR" "$@"

cp "$REPO_DIR/Passage/Passage.App.Linux/Assets/AppIcon.png" "$APP_DIR/passage.png"

mkdir -p "$BIN_DIR" "$(dirname "$DESKTOP_FILE")"
ln -sf "$APP_DIR/Passage.App.Linux" "$BIN_DIR/passage"

cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=Passage
Comment=Screenwriting software
Exec=$APP_DIR/Passage.App.Linux
Icon=$APP_DIR/passage.png
Terminal=false
Categories=Office;WordProcessor;
StartupWMClass=Passage.App.Linux
EOF

command -v update-desktop-database >/dev/null && update-desktop-database "$HOME/.local/share/applications" || true

echo
echo "Installed:"
echo "  App:      $APP_DIR"
echo "  Command:  passage  (via $BIN_DIR/passage)"
echo "  Launcher: $DESKTOP_FILE"
echo
echo "Re-run this script after code changes to update the installed app."
