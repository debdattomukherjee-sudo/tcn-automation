# TCN automation — working notes for Claude

## Git workflow (standing instruction)
At the **end of any session where files in this folder changed**, automatically
stage and commit the changes before wrapping up:

```bash
cd ~/Desktop/tcn-automation
git add -A
git commit -m "<concise summary of what changed this session>"
```

Then tell Dev it's committed and that the only remaining step on his side is
`git push`. Do **not** attempt to push yourself — pushing needs Dev's GitHub
credentials, which only exist on his Mac.

If nothing changed, skip the commit.

## Never commit secrets / state
These are gitignored and must stay local — never force-add them:
`credentials.json`, `token.json`, `slack_report_webhook.txt`,
`slack_alert_webhook.txt`, `history.json`, `processed.json`, `.venv/`,
`.venv-1/`, `_work/`, `__pycache__/`, `logs/`.

## Project shape (quick map)
- `run.py` — orchestrator (Drive poll → build → upload → Slack).
- `tcn_report.py` — Excel report builder (OB/IB tabs, Leadership View, Trends,
  Alerts, portfolio).
- `trends.py` — WoW/MoM/DoD history store + deltas (`history.json`).
- `alerts.py` — severity-scored alert engine on top of trends.
- `config.py` — clients, Slack, `ALERTS` thresholds, `TARGETS` SLAs.
- `slack_notify.py` — Workflow-webhook delivery (report + change digest).
- `drive_sync.py` — Google Drive API.
- `SETUP.md` — full setup/run/scheduling docs.
