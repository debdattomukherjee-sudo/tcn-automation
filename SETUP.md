# TCN Report Automation — VSCode Setup

This runs on your own machine. You drop a clean OB/IB workbook into a client's
Google Drive folder; this tool picks it up, builds the report, and uploads it
back into the same folder.

## What's in this folder
- `tcn_report.py` — the analysis engine (builds the 9-tab report).
- `drive_sync.py` — talks to Google Drive (login, download, upload).
- `run.py` — the thing you actually run; ties it all together.
- `config.py` — **the only file you edit** (clients + Drive folder IDs).
- `requirements.txt` — Python packages to install.

---

## Step 1 — Install Python
Install Python 3.10+ from https://www.python.org/downloads (macOS: `brew install python` also works).
Verify in a terminal:
```
python3 --version
```

## Step 2 — Open the project in VSCode
1. Put this whole folder somewhere stable, e.g. `~/tcn-automation`.
2. VSCode → File → Open Folder → pick that folder.
3. Install the **Python** extension (by Microsoft) if VSCode prompts you.

## Step 3 — Create a virtual environment and install packages
In the VSCode terminal (Terminal → New Terminal):
```
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
When VSCode asks "select interpreter", pick the one inside `.venv`.

## Step 4 — Get Google Drive credentials (one time, ~5 min)
This lets the script read/write your Drive.
1. Go to https://console.cloud.google.com → create a project (any name).
2. **APIs & Services → Library** → search "Google Drive API" → **Enable**.
3. **APIs & Services → OAuth consent screen** → choose **External** → fill the
   required name/email → Save. Under **Test users**, add your own email
   (debdatto.mukherjee@skit.ai).
4. **APIs & Services → Credentials → Create Credentials → OAuth client ID** →
   Application type: **Desktop app** → Create.
5. Click **Download JSON**. Rename the file to **`credentials.json`** and drop it
   into this project folder (next to `run.py`).

> Keep `credentials.json` and the `token.json` it generates private — don't commit
> them to git or share them.

## Step 5 — First run
```
python run.py
```
- A browser window opens → sign in → "Allow". This creates `token.json` so you
  won't be asked again.
- The script scans each client folder in `config.py`, processes any raw workbook,
  and uploads `<filename>_REPORT.xlsx` back to the same folder.

That's it. Open the report in Drive to check the numbers.

---

## Daily use
Drop your clean workbook (two tabs named **OB** and **IB**) into the client's
Drive folder, then run:
```
python run.py
```
Or leave it watching so you don't have to run it manually:
```
python run.py --watch
```
It re-checks every 5 minutes (change `POLL_INTERVAL_SECONDS` in `config.py`).

## File naming (drives client / period / dates automatically)
- Daily:   `PrimeRecovery_Daily_2026-06-24.xlsx`
- Weekly:  `PrimeRecovery_Weekly_2026-06-16_to_2026-06-22.xlsx`
- Monthly: `PrimeRecovery_Monthly_2026-06.xlsx`

The same code handles all three — daily (Tue–Fri), the weekly recap (Mon), and
the monthly recap. Just give the file the right name and date range.

## Adding another client
Open `config.py`, copy the Prime Recovery line, paste its Drive folder ID:
```python
CLIENTS = {
    "Prime Recovery": {"drive_folder_id": "1PH5PB0ScCV5u8HtKX8B5k_6SSIj9Niwd"},
    "Auto Finance":   {"drive_folder_id": "PASTE_FOLDER_ID_HERE"},
}
```

## Step 6 — Slack delivery (via Workflow Builder webhook — no app/admin approval)
Creating a Slack app needs admin approval, so instead we use **Workflow Builder**,
which any member can use. You build a workflow that starts "From a webhook" and
posts a message to you; Slack gives you a URL the script POSTs to.

> Heads-up: a webhook can only post a **text message**, so it cannot attach the
> Excel file. The message includes the summary plus a **Drive link** that opens
> the report in one click.

**Build the report workflow:**
1. In Slack: **Tools → Workflow Builder → Create → From a webhook**.
2. Click **Set Up Variables** → add ONE variable:
   - Key: `text`  · Data type: **Text**  → Done.
3. **Continue**, then **+ Add Step → Messages → "Send a message to a channel"**
   (or to a person — pick yourself for a DM).
4. In the message box, use the variable picker to insert **`text`** so the
   message body is just that variable. Save.
5. **Publish** the workflow. Open the webhook trigger and **Copy** the
   *Web request URL*.
6. In your project folder create **`slack_report_webhook.txt`** and paste the URL
   (just the URL, nothing else). Save.

**Optional — separate alerts:** repeat the steps for a second workflow (e.g.
posting to an "alerts" channel), and paste its URL into
**`slack_alert_webhook.txt`**. If you skip this, alerts go to the report webhook.

Now run again:
```
python run.py
```
You should get a Slack message with the summary and a Drive link. (Reports
already processed won't resend — drop a fresh file or clear its entry in
`processed.json` to re-test.)

> To switch Slack off temporarily, set `SLACK["enabled"] = False` in `config.py`.
> Keep the `slack_*_webhook.txt` files private — anyone with the URL can post.
> They're already in `.gitignore`.

---

## Cross-client portfolio roll-up (automatic — no setup)
When a single `run.py` processes **2+ clients**, it also builds a one-page
**`PORTFOLIO_<period>_<dates>_REPORT.xlsx`**, uploads it to the **parent** Drive
folder, and posts an exec digest to Slack. It ranks every client by
outbound connect rate, totals the portfolio, and flags the top performer and the
one needing attention. Nothing to configure — it just happens.

---

## Trend analysis — WoW / MoM / DoD (automatic — no setup)
Every report you build is recorded to **`history.json`** (kept in the project
folder, git-ignored). From the *second* comparable period onward, the tool
compares the current report to the right prior period and shows the raw move on
each metric — no significance filtering, every change is shown as-is.

Which comparison runs is driven by the file's period (from its name):
- **Daily** → day-over-day **and** vs the same weekday last week.
- **Weekly** → week-over-week.
- **Monthly** → month-over-month.

Where the trends show up:
- A new **Trends** tab in each client report: per-metric table with Prior /
  Current / Δ / %Δ, colored green (improved) or red (worsened), plus a flag when
  the best calling window shifted.
- The **Leadership View** gets a one-line connect-rate movement badge.
- The **Slack** summary leads with a connect move line, e.g.
  `📈 Week-over-week: connect ▲ +4.5 pts (30.0% → 34.5%)`.
- The **portfolio roll-up** annotates each client with its connect pts move +
  rank change vs last period and calls out the **biggest mover**.

The very first time you run a given client/period there's nothing to compare to,
so the Trends tab is skipped and Slack just shows the baseline numbers. It fills
in automatically on the next run of the same period type.

> Rates move in **percentage points** (e.g. 30.0% → 34.5% = +4.5 pts), while
> counts and dollars also show a relative **%Δ**. To reset history, delete
> `history.json`.

---

## Alerts — "what changed, and is it worth a look?" (automatic)
On top of the raw trends, every run scores the changes and flags the ones that
cross a threshold. Trends still show **every** move untouched; alerts just add a
severity badge so nobody has to eyeball the whole table.

Three severities: 🔴 **CRITICAL**, 🟠 **WARN**, 🔵 **INFO**.

What gets watched (21 checks across four tiers):
- **Client deltas & levels** — connect-rate drop or floor breach, cost/connect
  spike or ceiling, spend-up-while-connects-flat, dial-volume crash,
  dials/connect deterioration, inbound miss-rate spike or ceiling, miss
  concentration in one window, inbound call surge.
- **Trend shape** — losing streak, deviation from the rolling baseline, all-time
  record high/low, best-window drift.
- **Portfolio** — below the portfolio average, rank drop, new worst performer,
  blended portfolio connect-rate drop.
- **Targets & data quality** — SLA target misses, month-to-date spend pacing,
  and sanity checks (rate out of range, spend with zero dials, missing file).

Where alerts show up:
- A dedicated **Alerts** tab in each client report: severity-colored table with
  Prior / Current / Change / what it means.
- The **Leadership View** gets an "Alerts this run" block.
- A separate **leadership change-summary** is posted to the **same Slack
  channel**, **right after** the report message, on the **same cadence** as the
  report (daily → day-over-day, weekly → WoW, monthly → MoM). When nothing trips
  a threshold it still posts a one-line "✅ no alerts, biggest move was…".
- The **portfolio roll-up** posts its own exec alert digest when portfolio
  thresholds fire.

Tuning thresholds lives in **`config.py`**:
- `ALERTS = {...}` — every threshold (drop points, spike %, streak length,
  baseline window, etc.). Client-specific floors/ceilings default to `None`
  (off) so nothing false-fires out of the box; delta, streak, baseline,
  portfolio and data-quality checks work immediately.
- `TARGETS = {...}` — opt-in per-client SLAs (`connect_rate`, `max_miss_rate`,
  `monthly_spend`). Only clients you list here get target alerts.
- `slack_min_severity` controls the floor for what reaches Slack (the Excel tab
  always shows everything).

---

## Agents & dispositions — both streams (automatic — no setup)
Both dumps now carry agent + disposition columns: **Agent First Name**, **Agent
Last Name** and **An Agent Call Response** (the disposition), plus **Agent Call
Talk / Hold / Wrap-up Duration**. The engine reads them on **both outbound and
inbound** and adds a full agent + disposition layer on the same Analysis →
drill-down → trends → alerts pattern as everything else.

> Note on the OB dump: starting 2026-06-29 these extra columns (Col R onward) are
> added **manually** to every outbound dump. The engine matches them by header
> name, so column position doesn't matter — drop them anywhere to the right.

Disposition codes are read as structured `class-party-outcome` (e.g.
`C-RPC-PTP` = **C**onnected, **R**ight-**P**arty **C**ontact, **P**romise To
Pay). Anything starting with `C-` counts as **connected**. Rows with no agent
name are system calls (abandoned / voicemail / hang-up before pickup) — they're
bucketed as **"System / Unassigned"** and kept out of the per-agent rankings.

Four new tabs in each report — the same pair for each stream:
- **OB Agents** / **IB Agents** — one row per named agent: calls handled,
  connected, connect rate, PTP / Payment / No-PTP counts, average talk time and
  their most common disposition, ranked by volume then connect rate. The System /
  Unassigned bucket is shown separately at the bottom.
- **OB Dispositions** / **IB Dispositions** — connectivity summary (connected
  share of dispositioned calls), the full disposition distribution, and
  breakdowns by class (Connected vs Not), by party (RPC / TPC) and by outcome.

The Leadership View hub gains an **agents & dispositions** callout for each
stream (connected-disp rate, top / lowest agent, most common disposition).

Trends (from the second weekly/monthly period on):
- The **Trends** tab gains, **for each stream**, a **per-disposition** move table
  (share of dispositioned calls, Prior % → Current % → Δ, biggest shifts first)
  and a **per-agent** move table (connect rate Prior → Current → Δ, who went up /
  down).

Alerts (tuned in `config.py → ALERTS`; evaluated for **both** streams, tagged OB / IB):
- `connected_disp_drop_pts` (default 5) — 🟠 WARN if the connected (C-*) share of
  dispositioned calls falls by that many points.
- `agent_connect_move_pts` (default 8) — flags any agent whose connect rate moves
  that far: a **drop** is 🟠 WARN, a **rise** is 🔵 INFO. Guarded by
  `agent_min_calls` (default 20 in **both** periods, anti-noise) and capped at
  `agent_max_alerts`.
- `disposition_shift_pts` (default 5) — 🔵 INFO per disposition whose share moves
  that far, capped at `disposition_max_alerts`.

The first weekly dump has nothing to compare against, so you'll see the new tabs
populated but no agent/disposition trend or alert rows yet — those fill in on the
second weekly run automatically.

---

## Live formulas — every number is auditable (automatic — no setup)
Leadership can click any analytical cell in the report and see exactly how it was
derived. Instead of writing pre-computed numbers, the report embeds a copy of the
raw dump and writes every within-period calculation as a **visible Excel formula**
(`COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIFS`, and same-table cell
ratios). Open any cell and the formula bar shows the lookup.

Two extra tabs hold the source data the formulas point at:
- **OB_RawData** — the full outbound dump plus helper columns (Weekday, Week
  Block, Is Connect).
- **IB_RawData** — the full inbound dump plus the same helpers, and for
  dispositioned data: Agent Name, Disp Code, Has Disp, Disp Class / Party /
  Outcome, Disp Connected.

So a connect rate reads as `=Connects/Dials` (both live counts), `$ / connect` as
`=SUMIF(spend in window)/connects`, the disposition split as
`=COUNTIF(Disp Code, "C-RPC-PTP")`, and so on across OB, IB, agents and
dispositions. Edit a raw row and every dependent number recalculates in Excel.

What stays **computed** (by design): the cross-period **Trends** and **Alerts**
tabs (they compare against the prior period stored in `history.json`, which isn't
in the workbook), and the short argmax callouts on the Leadership page (e.g.
"Best day — Tuesday") that simply point you to the detail tab where the underlying
numbers are fully formula-driven.

---

## Run it automatically every day (macOS scheduling)
So you never have to run `run.py` by hand, install it as a daily LaunchAgent:
```
chmod +x install_schedule.sh uninstall_schedule.sh
./install_schedule.sh            # daily at 08:00
./install_schedule.sh 18 30      # or pick a time: 18:30
```
- Uses the project's `.venv` python automatically if present.
- Logs each run to `logs/tcn-report.out.log` / `.err.log`.
- Run it once immediately to test:
  `launchctl kickstart -k gui/$(id -u)/com.skit.tcn-report`
- Remove the schedule any time: `./uninstall_schedule.sh`

> The Mac must be powered on at the scheduled time. If it was asleep, macOS runs
> the job shortly after it wakes.

---

## Notes
- A file is reprocessed only if it's new or changed; reports (`*_REPORT.xlsx`) and
  temp files (`~$...`) are ignored as inputs.
- `processed.json` remembers what's been done. Delete a file's entry (or the whole
  file) to force a re-run.
- Delivery: Slack (summary + Drive link) fires per client, plus the portfolio
  roll-up across clients.
