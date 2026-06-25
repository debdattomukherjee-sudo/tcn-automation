#!/usr/bin/env python3
"""
Slack delivery via Workflow Builder webhooks (no Slack app / bot token / admin
approval needed).

How it works: you build a Slack Workflow that starts "From a webhook" with a
single Text variable named `text`, and a step that posts {{text}} as a message
to yourself or a channel. Slack gives you a URL; we POST JSON to it.

NOTE: Workflow webhooks can only post a text message - they cannot attach a
file. So the report's Drive link is included in the message instead of the
.xlsx. (file_path is accepted but only used to mention the file name.)

Webhook URLs are read from the files named in config.SLACK (or the
SLACK_REPORT_WEBHOOK / SLACK_ALERT_WEBHOOK env vars). If none is found, every
call is a no-op so the rest of the pipeline keeps working.
"""

import json
import os
import ssl
import urllib.request

import config

# Build an SSL context with a known-good CA bundle. On macOS, python.org builds
# often ship without system certs wired up, which makes HTTPS fail with
# CERTIFICATE_VERIFY_FAILED. certifi (a dependency of the Google libs) fixes it.
try:
    import certifi
    _SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except Exception:
    _SSL_CTX = ssl.create_default_context()


def _read_url(file_key, env_key):
    if os.environ.get(env_key):
        return os.environ[env_key].strip()
    path = config.SLACK.get(file_key)
    if path and os.path.exists(path):
        with open(path) as fh:
            return fh.read().strip()
    return None


def _post(url, text):
    """POST the message to a Slack workflow webhook. The JSON key must match the
    workflow's variable name exactly (case-sensitive) - set in config.SLACK."""
    var = config.SLACK.get("webhook_var", "Text")
    data = json.dumps({var: text}).encode("utf-8")
    req = urllib.request.Request(
        url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20, context=_SSL_CTX) as resp:
        return resp.status


def send_report(text, file_path=None, drive_link=None):
    """Post the report summary (+ Drive link) to the report webhook."""
    if not config.SLACK.get("enabled", True):
        return
    url = _read_url("report_webhook_file", "SLACK_REPORT_WEBHOOK")
    if not url:
        print("   (no Slack report webhook found - skipping Slack)")
        return
    body = text
    if drive_link:
        body += f"\n\n📎 Open report in Google Drive: {drive_link}"
    elif file_path:
        body += f"\n\n📎 Report file: {os.path.basename(file_path)}"
    try:
        _post(url, body)
        print("   Slack report sent via webhook")
    except Exception as e:
        print(f"   Slack report FAILED: {e}")


def send_report_channel(text):
    """Post a plain-text message to the SAME channel as the report (report
    webhook), no Drive-link appendage. Used for the leadership change-summary
    that follows each report message."""
    if not config.SLACK.get("enabled", True):
        return
    url = _read_url("report_webhook_file", "SLACK_REPORT_WEBHOOK")
    if not url:
        return  # no report webhook -> nothing to do (already warned on the report)
    try:
        _post(url, text)
        print("   Slack change-summary sent via webhook")
    except Exception as e:
        print(f"   Slack change-summary FAILED: {e}")


def send_alert(text):
    """Post an error/anomaly alert to the alert webhook (falls back to report)."""
    if not config.SLACK.get("enabled", True):
        return
    url = (_read_url("alert_webhook_file", "SLACK_ALERT_WEBHOOK")
           or _read_url("report_webhook_file", "SLACK_REPORT_WEBHOOK"))
    if not url:
        print("   (no Slack webhook found - skipping alert)")
        return
    try:
        _post(url, f":rotating_light: TCN automation alert\n{text}")
        print("   Slack alert sent via webhook")
    except Exception as e:
        print(f"   Could not send Slack alert: {e}")
