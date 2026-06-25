# =============================================================================
# TCN Report Automation - configuration
# Edit THIS file to onboard a new client or change settings. Nothing else.
# =============================================================================

# ----------------------------------------------------------------------------
# CLIENT FOLDERS - two ways to set this up. AUTO-DISCOVERY is the default.
# ----------------------------------------------------------------------------
#
# MODE 1 (recommended) - AUTO-DISCOVERY:
#   Put every client folder inside ONE parent folder in Drive. Set its ID below.
#   run.py treats each SUBFOLDER as a client (folder name = client name) and
#   processes them all. Add a new client later by just creating a new subfolder
#   in Drive - no code change needed.
#
#   To get the parent ID: open the parent folder in Drive and copy the last part
#   of the URL:  https://drive.google.com/drive/folders/<THIS_IS_THE_ID>
PARENT_FOLDER_ID = "1qcVRYiVVcISvrI_znAa1VUOWCVsfMm94"

# MODE 2 - EXPLICIT LIST (overrides auto-discovery if non-empty):
#   Leave PARENT_FOLDER_ID = None and list clients here instead. Useful if your
#   client folders are scattered in different places in Drive.
CLIENTS = {
    # "Prime Recovery": {"drive_folder_id": "1PH5PB0ScCV5u8HtKX8B5k_6SSIj9Niwd"},
    # "Auto Finance":   {"drive_folder_id": "PASTE_FOLDER_ID_HERE"},
}

# Subfolder names to ignore during auto-discovery (e.g. archives, scratch).
DISCOVERY_EXCLUDE = {"_archive", "archive", "old", "templates"}

# How often to re-check Drive for new files when running in --watch mode.
# 300 = every 5 minutes. (Google has no instant "file dropped" trigger, so we poll.)
POLL_INTERVAL_SECONDS = 300

# A file is treated as a generated REPORT (and skipped as an input) if its name
# contains this tag. Keep in sync with the report naming in run.py.
REPORT_TAG = "_REPORT"

# --- Slack delivery (via Workflow Builder webhooks - no app/admin approval) ---
# You build a Slack Workflow that starts "From a webhook" and posts a message to
# yourself (or a channel). It gives you a URL; paste it into the file below.
# The report goes to report_webhook; alerts go to alert_webhook (falls back to
# the report webhook if you only set one).
SLACK = {
    "report_webhook_file": "slack_report_webhook.txt",  # paste report workflow URL here
    "alert_webhook_file":  "slack_alert_webhook.txt",   # paste alert workflow URL here
    "webhook_var": "Text",  # MUST exactly match the variable key in your workflow (case-sensitive)
    "enabled": True,                                     # set False to skip Slack entirely
}

# --- Alerting (the "what changed, by how much, vs when" layer) ---------------
# Alerts ride on top of the trend deltas. The basis matches the report period:
#   daily file  -> day-over-day,  weekly -> week-over-week,  monthly -> MoM.
# A leadership change-summary is posted to the SAME Slack channel right AFTER
# the report message, and every fired alert is listed on the report's Alerts tab.
#
# Rates are compared in percentage POINTS (e.g. 30% -> 25% = -5 pts).
# Anything set to None is OFF (mostly client-specific floors/ceilings/targets).
ALERTS = {
    "enabled": True,

    # Outbound connect rate (percentage points vs prior period)
    "ob_connect_drop_warn_pts": 3.0,    # WARN if connect rate falls >= this
    "ob_connect_drop_crit_pts": 6.0,    # CRITICAL if it falls >= this
    "ob_connect_floor": None,           # e.g. 0.25 -> WARN below 25%. None = off

    # Cost efficiency
    "cost_per_connect_spike_pct": 0.15, # WARN if $/connect up >= 15% vs prior
    "cost_per_connect_ceiling": None,   # e.g. 0.60 -> WARN above $0.60. None = off
    "spend_up_pct": 0.10,               # spend up >= 10% AND connects flat/down -> WARN

    # Volume
    "dial_volume_crash_pct": 0.40,      # WARN if dials down >= 40% (stalled campaign)
    "dials_per_connect_rise_pct": 0.15, # WARN if dials needed per connect up >= 15%
    "ib_volume_surge_pct": 0.40,        # INFO if inbound calls up >= 40% (demand spike)

    # Inbound staffing
    "miss_rate_spike_pts": 5.0,         # WARN if miss rate up >= 5 pts
    "miss_rate_ceiling": 0.20,          # WARN if miss rate above 20%. None = off
    "miss_concentration_share": 0.40,   # WARN if >= 40% of misses fall in one window

    # Trend / pattern (use accumulated history.json)
    "streak_periods": 3,                # WARN if connect rate down N periods running
    "baseline_window": 4,               # trailing-N same-period average
    "baseline_deviation_pts": 3.0,      # WARN if current beyond +/- this vs baseline
    "record_lookback": 6,               # window for "best-ever / worst-in-N" (0 = all)

    # Best calling window drift
    "best_window_drift_alert": True,    # WARN when the best window moves to a new hour

    # Cross-client / portfolio
    "below_portfolio_pts": 5.0,         # WARN if a client is >= 5 pts under blended avg
    "rank_drop_spots": 2,               # WARN if a client falls >= 2 ranks vs prior
    "portfolio_blended_drop_pts": 2.0,  # WARN if the whole book's blended rate drops

    # Operational / data quality
    "data_quality_checks": True,        # bad values, zero-dial-with-spend, missing stream
    "expect_missing_file_alert": False, # WARN a client produced no file this run (opt-in)

    # Slack digest: minimum severity to include ("INFO" | "WARN" | "CRITICAL")
    "slack_min_severity": "INFO",
}

# Per-client goals (opt-in). Without an entry, target/pacing alerts are skipped.
#   connect_rate   -> WARN if outbound connect rate below it
#   max_miss_rate  -> WARN if inbound miss rate above it
#   monthly_spend  -> month-end spend pacing alert (monthly reports only)
TARGETS = {
    # "Prime Recovery": {"connect_rate": 0.30, "max_miss_rate": 0.15, "monthly_spend": 20000},
}

# --- Files kept inside this project folder (no need to change) ---------------
CREDENTIALS_FILE = "credentials.json"  # downloaded from Google Cloud (you create this)
TOKEN_FILE = "token.json"              # auto-created after first Google login
STATE_FILE = "processed.json"          # remembers which files were already done
HISTORY_FILE = "history.json"          # per-period metrics for WoW/MoM/DoD trends
WORK_DIR = "_work"                     # scratch space for downloads/outputs
