#!/usr/bin/env bash
# Remove the TCN report daily LaunchAgent.
set -euo pipefail

LABEL="com.skit.tcn-report"
UID_NUM="$(id -u)"
PLIST_DST="$HOME/Library/LaunchAgents/$LABEL.plist"

launchctl bootout "gui/$UID_NUM/$LABEL" 2>/dev/null || true
rm -f "$PLIST_DST"
echo "Uninstalled $LABEL. The schedule will no longer run."
