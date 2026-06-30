#!/usr/bin/env python3
"""
TCN Report Automation - orchestrator.

For every client in config.py:
  1. Look in its Drive folder for NEW raw OB/IB workbooks.
  2. Download, run the analysis engine (tcn_report.py).
  3. Upload the generated <name>_REPORT.xlsx back into the same folder.

Run modes:
    python run.py            # process anything new once, then exit
    python run.py --watch    # keep polling every POLL_INTERVAL_SECONDS
"""

import argparse
import json
import os
import time
import traceback
from datetime import datetime

import alerts
import config
import drive_sync
import slack_notify
import tcn_report
import trends


# --------------------------------------------------------------------------- #
# Processed-file memory (so we never redo the same file/version)               #
# --------------------------------------------------------------------------- #
def load_state():
    if os.path.exists(config.STATE_FILE):
        with open(config.STATE_FILE) as fh:
            return json.load(fh)
    return {}


def save_state(state):
    with open(config.STATE_FILE, "w") as fh:
        json.dump(state, fh, indent=2)


def is_raw_input(name, mime=None):
    """A file counts as a raw OB/IB input if it's either a real .xlsx/.xls
    upload OR a native Google Sheet (which we export to xlsx on download).
    Generated reports (REPORT_TAG) and Office temp files (~$) are excluded.
    Native Sheets are uploaded by us only as xlsx, so a Sheet is always raw."""
    if name.startswith("~$") or config.REPORT_TAG.lower() in name.lower():
        return False
    if mime == drive_sync.GSHEET_MIME:
        return True
    return name.lower().endswith((".xlsx", ".xls"))


def input_base_name(name):
    """Filename minus a trailing .xlsx/.xls only. Native Sheet names have no
    extension (and may legitimately contain dots like '6.29'), so we must NOT
    naively split on the last dot."""
    low = name.lower()
    for ext in (".xlsx", ".xls"):
        if low.endswith(ext):
            return name[: -len(ext)]
    return name


def period_bounds(ob, ib):
    """Earliest / latest EST date across both streams (anchors trend history)."""
    dts = []
    for df in (ob, ib):
        if df is not None:
            dts += [d for d in df["EST Date"].dropna().tolist()]
    if not dts:
        return None, None
    return min(dts), max(dts)


def date_label_from_data(ob, ib, from_filename):
    if from_filename:
        return from_filename
    dts = []
    for df in (ob, ib):
        if df is not None:
            dts += [d for d in df["EST Date"].dropna().tolist()]
    if dts:
        lo, hi = min(dts), max(dts)
        return str(lo) if lo == hi else f"{lo} to {hi}"
    return None


def _period_lines(res, period):
    """Extra Slack lines for weekly (best/worst day) and monthly (best/worst
    week + best day overall + best day per week). Ranked by connect rate.
    Plain text only (Workflow webhook doesn't render mrkdwn), so we lead with
    emoji markers for hierarchy."""
    out = []
    if period == "weekly":
        bd = tcn_report.best_by_rate(res.get("dow"))
        wd = tcn_report.best_by_rate(res.get("dow"), want="worst")
        if bd:
            out.append(f"     🏆 Best day: {bd[0]} — {bd[3]:.1%} ({bd[1]}/{bd[2]})")
        if wd and (not bd or wd[0] != bd[0]):
            out.append(f"     🔻 Worst day: {wd[0]} — {wd[3]:.1%} ({wd[1]}/{wd[2]})")
    elif period == "monthly":
        bw = tcn_report.best_by_rate(res.get("weekblock"))
        ww = tcn_report.best_by_rate(res.get("weekblock"), want="worst")
        if bw:
            out.append(f"     📅 Best week: {bw[0]} — {bw[3]:.1%} ({bw[1]}/{bw[2]})")
        if ww and (not bw or ww[0] != bw[0]):
            out.append(f"     📅 Worst week: {ww[0]} — {ww[3]:.1%} ({ww[1]}/{ww[2]})")
        bd = tcn_report.best_by_rate(res.get("dow"))
        if bd:
            out.append(f"     🏆 Best day overall: {bd[0]} — {bd[3]:.1%} ({bd[1]}/{bd[2]})")
        block_dow = res.get("block_dow", {})
        if block_dow:
            parts = []
            for i, b in enumerate(tcn_report.WEEKBLOCK_ORDER, start=1):
                day = tcn_report.best_by_rate(block_dow.get(b))
                if day:
                    parts.append(f"W{i} {day[0][:3]}")
            if parts:
                out.append("     🗓️ Best day each week: " + ", ".join(parts))
    return out


def _agent_disp_lines(res, label):
    """Slack lines for one stream's agent + disposition layer: dispositioned
    volume, connected (C-*) rate, RPC PTPs + payments, and the top agents ranked
    by PTP + payment. Empty when the dump carried no disposition column."""
    ad = res.get("agentdisp") if res else None
    if not ad or not ad.get("disp_total"):
        return []
    # Use the engine's canonical totals (taxonomy-aware) rather than matching
    # outcome token strings, which vary (e.g. PAYMENT_ON_CALL, not PAYMENT).
    ptp = int(ad.get("ptp_total", 0))
    pay = int(ad.get("payment_total", 0))
    out = ["",
           f"🧑‍💼  {label} AGENTS & DISPOSITIONS",
           f"     {ad['disp_total']:,} dispositioned  ·  "
           f"{ad['disp_connected_rate']:.1%} connected (C-*)",
           f"     🤝 RPC PTPs: {ptp:,}  ·  💵 Payments: {pay:,}"]
    agents = ad.get("agents") or []
    ranked = sorted(agents, key=lambda a: (a.get("ptp", 0) + a.get("payment", 0)),
                    reverse=True)
    top = [a for a in ranked if (a.get("ptp", 0) + a.get("payment", 0)) > 0][:3]
    if top:
        out.append("     🏅 Top agents (PTP / payment):")
        for a in top:
            out.append(f"        • {a['agent']}: {a['ptp']} PTP / {a['payment']} pay "
                       f"({a['rate']:.0%} conn, {a['handled']:,} calls)")
    return out


def build_summary(client, period, date_label, ob_res, ib_res, comparisons=None):
    """Clean plain-text summary for Slack. Workflow Builder posts the variable
    as-is and does NOT render mrkdwn, so we avoid *bold* markup and lean on
    emoji section markers + indentation for hierarchy. Weekly/monthly add
    day/week breakdowns."""
    # Leading blank line keeps this off the workflow's "New TCN Report" header.
    lines = ["",
             f"📊  {client} — {period.title()} ({date_label})",
             "━━━━━━━━━━━━━━━━━━━━"]
    mv = trends.primary_connect_move(comparisons) if comparisons else None
    if mv:
        lines.append(f"     📈 {mv['basis']}: connect {mv['arrow']} {mv['delta_str']} "
                     f"({mv['prior']:.1%} → {mv['current']:.1%})")
    if ob_res:
        lines.append("")
        lines.append("📞  OUTBOUND")
        lines.append(
            f"     {ob_res['total']:,} dials  ·  "
            f"{ob_res['answered_linkcalls']} answered linkcalls  ·  "
            f"{ob_res['connectivity']:.1%} connect")
        # Best window ranked by CONNECT RATE (matches the Excel 'Best Time to Call'),
        # not by call volume. peak_windows applies a volume floor so it's not a fluke.
        peaks = tcn_report.peak_windows(ob_res.get("rate_est"))
        if peaks:
            w, _s, _t, r = peaks[0]
            lines.append(f"     🕐 Best window to call (EST): {w} — {r:.0%} connect")
        else:
            be, ne, pe = tcn_report.best_bucket(ob_res["est_al"])
            lines.append(f"     🕐 Best window (EST): {be} ({pe:.0%})")
        lines += _period_lines(ob_res, period)
        if ob_res.get("total_cost") is not None:
            lines.append(f"     💰 ${ob_res['total_cost']:,.2f} total  ·  "
                         f"${ob_res.get('cost_per_al', 0):,.3f} / linkcall")
        lines += _agent_disp_lines(ob_res, "OB")
    if ib_res:
        be, ne, pe = tcn_report.best_bucket(ib_res["est_conn"])
        lines.append("")
        lines.append("📥  INBOUND")
        lines.append(
            f"     {ib_res['total']:,} calls  ·  "
            f"{ib_res['connected']} connected  ·  "
            f"{ib_res['connectivity']:.1%}  ·  {ib_res['miss_total']} missed")
        # Inbound is demand-driven, so we lead with the busiest window (when calls
        # actually arrive) and flag where they go unanswered — the staffing signal.
        lines.append(f"     🕐 Busiest window (EST): {be} ({pe:.0%})")
        mb = tcn_report.best_bucket(ib_res["miss_est"])
        if mb[1]:
            lines.append(f"     🚨 Most-missed window (EST): {mb[0]} ({mb[1]} missed)")
        lines += _period_lines(ib_res, period)
        lines += _agent_disp_lines(ib_res, "IB")
    return "\n".join(lines)


def process_file(service, client, folder_id, fobj, state, hist):
    name = fobj["name"]
    key = f"{fobj['id']}::{fobj['modifiedTime']}"
    if state.get(key):
        return None  # this exact version already done

    print(f"[{client}] New file detected: {name}")
    os.makedirs(config.WORK_DIR, exist_ok=True)
    base = input_base_name(name)
    # Always land the downloaded/exported bytes as a real .xlsx so pandas picks
    # the openpyxl engine regardless of the source name (native Sheets carry no
    # extension; download_xlsx exports them to xlsx).
    local_in = os.path.join(config.WORK_DIR, base + ".xlsx")
    drive_sync.download_xlsx(service, fobj, local_in)

    _, file_period, file_date = tcn_report.parse_filename(name)
    period = file_period or "daily"

    ob, ib = tcn_report.load(local_in)
    if ob is None and ib is None:
        print(f"[{client}] SKIP: no OB/IB tabs found in {name}")
        return None

    ob_res = tcn_report.analyze_ob(ob) if ob is not None else None
    ib_res = tcn_report.analyze_ib(ib) if ib is not None else None
    date_label = date_label_from_data(ob, ib, file_date) or "-"
    p_start, p_end = period_bounds(ob, ib)

    # Build this run's record and compare it to the prior period of the same
    # type BEFORE we add it to history (so it never compares against itself).
    rec = tcn_report.client_metrics(client, period, date_label, ob_res, ib_res,
                                    drive_link=None,
                                    period_start=p_start, period_end=p_end)
    comparisons = trends.compute_comparisons(rec, hist)
    client_alerts = alerts.evaluate(rec, comparisons, hist)

    out_name = base + config.REPORT_TAG + ".xlsx"
    local_out = os.path.join(config.WORK_DIR, out_name)
    tcn_report.build_workbook(client, period, date_label, ob_res, ib_res, local_out,
                              comparisons=comparisons, alerts=client_alerts)

    uploaded = drive_sync.upload_xlsx(service, local_out, folder_id, out_name)
    drive_link = uploaded.get("webViewLink")
    rec["drive_link"] = drive_link
    print(f"[{client}] Report uploaded -> {out_name}")
    if ob_res:
        print(f"           OB: {ob_res['total']} dials, {ob_res['answered_linkcalls']} "
              f"answered linkcalls ({ob_res['connectivity']:.1%})")
    if ib_res:
        print(f"           IB: {ib_res['total']} inbound, {ib_res['connected']} connected "
              f"({ib_res['connectivity']:.1%}), {ib_res['miss_total']} missed")

    summary = build_summary(client, period, date_label, ob_res, ib_res, comparisons)
    slack_notify.send_report(summary, file_path=local_out, drive_link=drive_link)

    # Leadership change-summary — posted to the SAME channel, right AFTER the
    # report message, on the report's cadence (daily=DoD, weekly=WoW, monthly=MoM).
    digest = alerts.client_digest(client, period, client_alerts,
                                  trends.primary_connect_move(comparisons))
    slack_notify.send_report_channel(digest)

    state[key] = {"client": client, "report": out_name,
                  "processed_at": datetime.now().isoformat()}
    save_state(state)

    # Persist this period to history so the next comparable period can trend it.
    trends.append_record(hist, rec)
    trends.save_history(hist)
    return rec


def process_client(service, client, folder_id, state, hist):
    """Process every new raw file for a client. Returns the list of metric
    records produced this run (one per newly-built report) for the roll-up."""
    records = []
    for fobj in drive_sync.list_files(service, folder_id):
        if not is_raw_input(fobj["name"], fobj.get("mimeType")):
            continue
        try:
            rec = process_file(service, client, folder_id, fobj, state, hist)
            if rec:
                records.append(rec)
        except Exception:
            print(f"[{client}] ERROR while processing {fobj['name']}:")
            tb = traceback.format_exc()
            print(tb)
            slack_notify.send_alert(
                f"*{client}* failed while processing `{fobj['name']}`\n```{tb[-1500:]}```")
    return records


def discover_clients(service):
    """Build the {client_name: folder_id} map.

    If config.CLIENTS is populated it wins (explicit override). Otherwise every
    subfolder under config.PARENT_FOLDER_ID becomes a client, with the folder
    name as the client name. New client = new subfolder in Drive, no code edit."""
    if config.CLIENTS:
        return {name: cfg["drive_folder_id"]
                for name, cfg in config.CLIENTS.items()}

    parent = getattr(config, "PARENT_FOLDER_ID", None)
    if not parent:
        print("No PARENT_FOLDER_ID and no CLIENTS configured - nothing to do.")
        return {}

    exclude = {x.lower() for x in getattr(config, "DISCOVERY_EXCLUDE", set())}
    clients = {}
    for f in drive_sync.list_subfolders(service, parent):
        if f["name"].lower() in exclude:
            continue
        clients[f["name"]] = f["id"]
    print(f"Discovered {len(clients)} client folder(s): "
          f"{', '.join(sorted(clients)) or '(none)'}")
    return clients


def build_rollup_summary(records, period, date_label, movement=None):
    """Plain-text exec digest for Slack ranking clients by connect rate.
    When movement (prior-period deltas) is available, annotate each client with
    its connect-rate pts move + rank change, and call out the biggest mover."""
    basis = {"weekly": "WoW", "monthly": "MoM"}.get(period, "DoD")
    ob_recs = [r for r in records if r.get("ob")]
    lines = ["",
             f"📈  PORTFOLIO ROLL-UP — {period.title()} ({date_label})",
             "━━━━━━━━━━━━━━━━━━━━"]
    if ob_recs:
        td = sum(r["ob"]["dials"] for r in ob_recs)
        tc = sum(r["ob"]["connects"] for r in ob_recs)
        rate = (tc / td) if td else 0
        lines.append(f"     {len(ob_recs)} clients · {td:,} dials · "
                     f"{tc:,} connects · {rate:.1%} blended connect")
        lines.append("")
        lines.append("📞  Outbound by client (best first)")
        for r in sorted(ob_recs, key=lambda x: x["ob"]["rate"], reverse=True):
            o = r["ob"]
            line = (f"     • {r['client']}: {o['rate']:.1%} "
                    f"({o['connects']:,}/{o['dials']:,})")
            mv = (movement or {}).get(r["client"])
            if mv and mv.get("delta_pts") is not None:
                arrow = "▲" if mv["delta_pts"] > 0 else ("▼" if mv["delta_pts"] < 0 else "▬")
                line += f"  [{arrow} {abs(mv['delta_pts']) * 100:.1f} pts {basis}]"
            lines.append(line)
        # Biggest mover by absolute connect-rate change.
        movers = [(c, m) for c, m in (movement or {}).items()
                  if m.get("delta_pts") is not None]
        if movers:
            c, m = max(movers, key=lambda kv: abs(kv[1]["delta_pts"]))
            arrow = "▲" if m["delta_pts"] > 0 else ("▼" if m["delta_pts"] < 0 else "▬")
            lines.append("")
            lines.append(f"     🚀 Biggest mover ({basis}): {c} "
                         f"{arrow} {abs(m['delta_pts']) * 100:.1f} pts")
    ib_recs = [r for r in records if r.get("ib")]
    if ib_recs:
        lines.append("")
        lines.append("📥  Inbound miss rate by client")
        for r in sorted(ib_recs, key=lambda x: x["ib"]["miss_rate"]):
            b = r["ib"]
            lines.append(f"     • {r['client']}: {b['miss_rate']:.1%} missed "
                         f"({b['missed']:,}/{b['calls']:,})")

    # Client-wise PTPs & payments (across OB + IB dispositioned calls).
    def _client_ptp(r):
        ptp = pay = 0
        for stream in ("ob", "ib"):
            s = r.get(stream) or {}
            for a in (s.get("agents") or {}).values():
                ptp += a.get("ptp", 0)
                pay += a.get("payment", 0)
        return ptp, pay
    ptp_rows = [(r["client"], *_client_ptp(r)) for r in records]
    ptp_rows = [x for x in ptp_rows if x[1] or x[2]]
    if ptp_rows:
        tot_ptp = sum(x[1] for x in ptp_rows)
        tot_pay = sum(x[2] for x in ptp_rows)
        lines.append("")
        lines.append(f"🤝  PTPs & payments by client  (total {tot_ptp:,} PTP · {tot_pay:,} pay)")
        for c, ptp, pay in sorted(ptp_rows, key=lambda x: -(x[1] + x[2])):
            lines.append(f"     • {c}: {ptp:,} PTP · {pay:,} pay")
    return "\n".join(lines)


def publish_rollup(service, records, hist):
    """Build the portfolio workbook, upload it to the parent folder, and send
    the exec digest to Slack. Skips quietly if <2 clients ran or no
    parent folder is configured."""
    if len(records) < 2:
        return  # roll-up only makes sense across multiple clients
    period = records[0].get("period", "daily")
    date_label = records[0].get("date_label", "-")

    # Per-client prior-period connect-rate move + rank change vs last period.
    movement = trends.rollup_movement(records, hist)
    port_alerts = alerts.evaluate_portfolio(records, movement, hist)

    os.makedirs(config.WORK_DIR, exist_ok=True)
    out_name = f"PORTFOLIO_{period.title()}_{date_label.replace(' ', '')}_REPORT.xlsx"
    local_out = os.path.join(config.WORK_DIR, out_name)
    tcn_report.build_portfolio(records, period, date_label, local_out,
                               movement=movement, alerts=port_alerts)
    print(f"[PORTFOLIO] Roll-up built: {out_name} ({len(records)} clients)")

    drive_link = None
    parent = getattr(config, "PARENT_FOLDER_ID", None)
    if parent:
        try:
            uploaded = drive_sync.upload_xlsx(service, local_out, parent, out_name)
            drive_link = uploaded.get("webViewLink")
            print(f"[PORTFOLIO] Uploaded to parent folder")
        except Exception:
            print("[PORTFOLIO] ERROR uploading roll-up:")
            print(traceback.format_exc())

    summary = build_rollup_summary(records, period, date_label, movement=movement)
    slack_notify.send_report(summary, file_path=local_out, drive_link=drive_link)

    # Portfolio alert digest — posted right after the roll-up message.
    pdigest = alerts.portfolio_digest(period, port_alerts)
    if pdigest:
        slack_notify.send_report_channel(pdigest)


def run_once():
    service = drive_sync.get_service()
    state = load_state()
    hist = trends.load_history()
    clients = discover_clients(service)
    all_records = []
    for client, folder_id in clients.items():
        try:
            all_records += process_client(service, client, folder_id, state, hist)
        except Exception:
            print(f"[{client}] ERROR listing folder:")
            tb = traceback.format_exc()
            print(tb)
            slack_notify.send_alert(
                f"*{client}* could not read its Drive folder\n```{tb[-1500:]}```")
    try:
        publish_rollup(service, all_records, hist)
    except Exception:
        print("[PORTFOLIO] ERROR building roll-up:")
        print(traceback.format_exc())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--watch", action="store_true",
                    help="poll Drive continuously instead of running once")
    args = ap.parse_args()

    if args.watch:
        print(f"Watching Drive every {config.POLL_INTERVAL_SECONDS}s. Press Ctrl+C to stop.")
        while True:
            run_once()
            time.sleep(config.POLL_INTERVAL_SECONDS)
    else:
        run_once()
        print("Done.")


if __name__ == "__main__":
    main()
