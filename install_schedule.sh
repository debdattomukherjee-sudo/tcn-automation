#!/usr/bin/env bash
# =============================================================================
# Install the TCN report as a daily macOS LaunchAgent (no manual run.py needed).
#
#   ./install_schedule.sh            # run daily at 08:00
#   ./install_schedule.sh 18 30      # run daily at 18:30 (24-hour clock)
#
# Re-run any time to change the time — it reloads cleanly. Uninstall with
# ./uninstall_schedule.sh
# =============================================================================
set -euo pipefail

LABEL="com.skit.tcn-report"
HOUR="${1:-8}"
MINUTE="${2:-0}"

# Project dir = where this script lives.
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PLIST_SRC="$PROJECT_DIR/com.skit.tcn-report.plist.template"
LA_DIR="$HOME/Library/LaunchAgents"
PLIST_DST="$LA_DIR/$LABEL.plist"

# Prefer the project venv python if present, else whatever python3 is on PATH.
if [ -x "$PROJECT_DIR/.venv/bin/python" ]; then
    PYTHON="$PROJECT_DIR/.venv/bin/python"
elif [ -x "$PROJECT_DIR/venv/bin/python" ]; then
    PYTHON="$PROJECT_DIR/venv/bin/python"
else
    PYTHON="$(command -v python3)"
fi

if [ ! -f "$PLIST_SRC" ]; then
    echo "ERROR: $PLIST_SRC not found. Run this from the project folder." >&2
    exit 1
fi

echo "Project : $PROJECT_DIR"
echo "Python  : $PYTHON"
echo "Schedule: daily at $(printf '%02d:%02d' "$HOUR" "$MINUTE")"

mkdir -p "$LA_DIR" "$PROJECT_DIR/logs"

# Fill the template placeholders.
sed -e "s#__PYTHON__#$PYTHON#g" \
    -e "s#__PROJECT_DIR__#$PROJECT_DIR#g" \
    -e "s#__HOUR__#$HOUR#g" \
    -e "s#__MINUTE__#$MINUTE#g" \
    "$PLIST_SRC" > "$PLIST_DST"

# Reload (bootout is harmless if it wasn't loaded).
UID_NUM="$(id -u)"
launchctl bootout "gui/$UID_NUM/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$UID_NUM" "$PLIST_DST"
launchctl enable "gui/$UID_NUM/$LABEL"

echo
echo "Installed. Verify with:  launchctl print gui/$UID_NUM/$LABEL | grep -i state"
echo "Run it now once with:    launchctl kickstart -k gui/$UID_NUM/$LABEL"
echo "Logs:                    $PROJECT_DIR/logs/tcn-report.{out,err}.log"
