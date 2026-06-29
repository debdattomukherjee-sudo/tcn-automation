#!/usr/bin/env python3
"""
Alert layer — "what changed, by how much, vs when, and is it worth attention?"

Alerts ride on top of the raw trend deltas (trends.py). They DON'T hide or
filter the deltas shown elsewhere; they add a severity flag and a plain-English
reason, routed to the report's Alerts tab + a leadership change-summary in Slack.

Basis matches the report period (driven by the file name):
    daily   -> day-over-day
    weekly  -> week-over-week
    monthly -> month-over-month

Two entry points:
    evaluate(rec, comparisons, hist)            -> per-client alerts
    evaluate_portfolio(records, movement, hist)  -> cross-client alerts

Each alert is a dict:
    {id, severity, stream, metric, basis, message,
     prior_str, current_str, change_str, window}
severity in {"CRITICAL","WARN","INFO"}; stream in {"OB","IB","Portfolio","Data"}.
"""

import config
import trends

SEV_RANK = {"CRITICAL": 0, "WARN": 1, "INFO": 2}
SEV_EMOJI = {"CRITICAL": "🔴", "WARN": "🟠", "INFO": "🔵"}


# --------------------------------------------------------------------------- #
# Config / formatting helpers                                                  #
# --------------------------------------------------------------------------- #
def _cfg():
    return getattr(config, "ALERTS", {}) or {}


def _targets():
    return getattr(config, "TARGETS", {}) or {}


def _fmtv(kind, v):
    if v is None:
        return "—"
    if kind == "rate":
        return f"{v * 100:.1f}%"
    if kind == "money":
        return f"${v:,.3f}"
    if kind == "count":
        return f"{v:,.0f}"
    return str(v)


def _pts(delta):
    sign = "+" if delta >= 0 else "−"
    return f"{sign}{abs(delta) * 100:.1f} pts"


def _pct(delta_frac):
    sign = "+" if delta_frac >= 0 else "−"
    return f"{sign}{abs(delta_frac) * 100:.1f}%"


def _a(aid, severity, stream, metric, basis, message,
       prior=None, current=None, kind=None, change_str="—", window=None):
    return {
        "id": aid, "severity": severity, "stream": stream, "metric": metric,
        "basis": basis, "message": message,
        "prior_str": _fmtv(kind, prior) if kind else "—",
        "current_str": _fmtv(kind, current) if kind else "—",
        "change_str": change_str, "window": window or "—",
    }


def _row(block, stream, metric):
    if not block:
        return None
    for r in block.get(stream, []):
        if r["metric"] == metric:
            return r
    return None


def _sorted(alerts):
    return sorted(alerts, key=lambda a: (SEV_RANK.get(a["severity"], 9), a["stream"]))


# --------------------------------------------------------------------------- #
# History series (for streak / baseline / record alerts)                       #
# --------------------------------------------------------------------------- #
def _prior_rates(hist, client, period):
    """Prior same-period outbound connect rates, oldest -> newest (current
    record is NOT in history yet, so this is purely the past)."""
    recs = hist.get(client, {}).get(period, [])
    return [r["ob"]["rate"] for r in recs if r.get("ob") and r["ob"].get("rate") is not None]


# --------------------------------------------------------------------------- #
# Per-client: deltas vs the primary comparison block                           #
# --------------------------------------------------------------------------- #
def _delta_alerts(blk, basis, cfg):
    out = []
    # ---- Outbound connect rate drop ----
    r = _row(blk, "ob", "Connect rate")
    if r and r["delta"] < 0:
        drop = -r["delta"] * 100
        warn = cfg.get("ob_connect_drop_warn_pts", 3.0)
        crit = cfg.get("ob_connect_drop_crit_pts", 6.0)
        sev = "CRITICAL" if drop >= crit else ("WARN" if drop >= warn else None)
        if sev:
            out.append(_a("ob_connect_drop", sev, "OB", "Connect rate", basis,
                          f"Connect rate fell {drop:.1f} pts {basis} "
                          f"({r['prior'] * 100:.1f}% → {r['current'] * 100:.1f}%)",
                          prior=r["prior"], current=r["current"], kind="rate",
                          change_str=_pts(r["delta"])))
    # ---- $/connect spike ----
    r = _row(blk, "ob", "$ / connect")
    if r and r["prior"]:
        p = (r["current"] - r["prior"]) / r["prior"]
        if p >= cfg.get("cost_per_connect_spike_pct", 0.15):
            out.append(_a("cost_per_connect_spike", "WARN", "OB", "$ / connect", basis,
                          f"$/connect up {p * 100:.1f}% {basis} "
                          f"(${r['prior']:.3f} → ${r['current']:.3f})",
                          prior=r["prior"], current=r["current"], kind="money",
                          change_str=_pct(p)))
    # ---- Dial-volume crash ----
    r = _row(blk, "ob", "Dials")
    if r and r["prior"]:
        p = (r["current"] - r["prior"]) / r["prior"]
        if p <= -cfg.get("dial_volume_crash_pct", 0.40):
            out.append(_a("dial_volume_crash", "WARN", "OB", "Dials", basis,
                          f"Dials down {abs(p) * 100:.1f}% {basis} "
                          f"({r['prior']:,.0f} → {r['current']:,.0f}) — campaign may be stalled",
                          prior=r["prior"], current=r["current"], kind="count",
                          change_str=_pct(p)))
    # ---- Dials/connect deteriorating ----
    r = _row(blk, "ob", "Dials / connect")
    if r and r["prior"]:
        p = (r["current"] - r["prior"]) / r["prior"]
        if p >= cfg.get("dials_per_connect_rise_pct", 0.15):
            out.append(_a("dials_per_connect_rise", "WARN", "OB", "Dials / connect", basis,
                          f"Dials per connect up {p * 100:.1f}% {basis} "
                          f"({r['prior']:.1f} → {r['current']:.1f}) — efficiency dropping",
                          prior=r["prior"], current=r["current"], kind="count",
                          change_str=_pct(p)))
    # ---- Spend up while connects flat/down ----
    rs = _row(blk, "ob", "Total spend")
    rc = _row(blk, "ob", "Connects")
    if rs and rc and rs["prior"]:
        sp = (rs["current"] - rs["prior"]) / rs["prior"]
        if sp >= cfg.get("spend_up_pct", 0.10) and rc["delta"] <= 0:
            out.append(_a("spend_up_connects_flat", "WARN", "OB", "Spend vs connects", basis,
                          f"Spend up {sp * 100:.1f}% but connects flat/down "
                          f"({rc['prior']:,.0f} → {rc['current']:,.0f}) — efficiency degraded",
                          prior=rs["prior"], current=rs["current"], kind="money",
                          change_str=_pct(sp)))
    # ---- Inbound miss-rate spike ----
    r = _row(blk, "ib", "Miss rate")
    if r and r["delta"] > 0 and r["delta"] * 100 >= cfg.get("miss_rate_spike_pts", 5.0):
        out.append(_a("miss_rate_spike", "WARN", "IB", "Miss rate", basis,
                      f"Miss rate up {r['delta'] * 100:.1f} pts {basis} "
                      f"({r['prior'] * 100:.1f}% → {r['current'] * 100:.1f}%) — staffing gap",
                      prior=r["prior"], current=r["current"], kind="rate",
                      change_str=_pts(r["delta"])))
    # ---- Inbound volume surge ----
    r = _row(blk, "ib", "Calls")
    if r and r["prior"]:
        p = (r["current"] - r["prior"]) / r["prior"]
        if p >= cfg.get("ib_volume_surge_pct", 0.40):
            out.append(_a("ib_volume_surge", "INFO", "IB", "Calls", basis,
                          f"Inbound calls up {p * 100:.1f}% {basis} "
                          f"({r['prior']:,.0f} → {r['current']:,.0f}) — demand spike, check coverage",
                          prior=r["prior"], current=r["current"], kind="count",
                          change_str=_pct(p)))
    return out


# --------------------------------------------------------------------------- #
# Per-client: absolute levels (floors / ceilings / concentration)              #
# --------------------------------------------------------------------------- #
def _level_alerts(rec, cfg):
    out = []
    ob, ib = rec.get("ob"), rec.get("ib")
    if ob:
        floor = cfg.get("ob_connect_floor")
        if floor is not None and ob.get("rate") is not None and ob["rate"] < floor:
            out.append(_a("ob_connect_floor", "WARN", "OB", "Connect rate", "vs floor",
                          f"Connect rate {ob['rate'] * 100:.1f}% is below the "
                          f"{floor * 100:.0f}% floor",
                          current=ob["rate"], kind="rate"))
        ceil = cfg.get("cost_per_connect_ceiling")
        cpc = ob.get("cost_per_connect")
        if ceil is not None and cpc is not None and cpc > ceil:
            out.append(_a("cost_ceiling", "WARN", "OB", "$ / connect", "vs ceiling",
                          f"$/connect ${cpc:.3f} is above the ${ceil:.3f} ceiling",
                          current=cpc, kind="money"))
    if ib:
        ceil = cfg.get("miss_rate_ceiling")
        mr = ib.get("miss_rate")
        if ceil is not None and mr is not None and mr > ceil:
            out.append(_a("miss_rate_ceiling", "WARN", "IB", "Miss rate", "vs ceiling",
                          f"Miss rate {mr * 100:.1f}% is above the {ceil * 100:.0f}% ceiling",
                          current=mr, kind="rate"))
        share = cfg.get("miss_concentration_share", 0.40)
        missed = ib.get("missed") or 0
        mmc = ib.get("most_missed_count") or 0
        if missed and share and (mmc / missed) >= share:
            out.append(_a("miss_concentration", "WARN", "IB", "Missed calls",
                          "within period",
                          f"{mmc / missed * 100:.0f}% of missed calls fall in "
                          f"{ib.get('most_missed_window')} — staff that window",
                          window=ib.get("most_missed_window")))
    return out


# --------------------------------------------------------------------------- #
# Per-client: trend / pattern (use accumulated history)                        #
# --------------------------------------------------------------------------- #
def _streak_alert(rec, hist, cfg):
    if not rec.get("ob"):
        return []
    n = int(cfg.get("streak_periods", 3) or 0)
    if n < 2:
        return []
    series = _prior_rates(hist, rec["client"], rec.get("period", "daily"))
    series = series + [rec["ob"]["rate"]]
    if len(series) < n + 1:
        return []
    tail = series[-(n + 1):]
    if all(tail[i + 1] < tail[i] for i in range(len(tail) - 1)):
        drop = (tail[0] - tail[-1]) * 100
        return [_a("connect_streak_down", "WARN", "OB", "Connect rate", f"{n}-period streak",
                   f"Connect rate has fallen {n} periods running "
                   f"({tail[0] * 100:.1f}% → {tail[-1] * 100:.1f}%, −{drop:.1f} pts) — sustained decline",
                   prior=tail[0], current=tail[-1], kind="rate", change_str=_pts(-drop / 100))]
    return []


def _baseline_alert(rec, hist, cfg):
    if not rec.get("ob"):
        return []
    win = int(cfg.get("baseline_window", 4) or 0)
    dev = cfg.get("baseline_deviation_pts", 3.0)
    series = _prior_rates(hist, rec["client"], rec.get("period", "daily"))
    if len(series) < 2 or win < 2:
        return []
    base = series[-win:]
    avg = sum(base) / len(base)
    cur = rec["ob"]["rate"]
    d_pts = (cur - avg) * 100
    if abs(d_pts) < dev:
        return []
    basis = f"vs trailing-{len(base)} avg"
    if d_pts < 0:
        return [_a("baseline_below", "WARN", "OB", "Connect rate", basis,
                   f"Connect rate {cur * 100:.1f}% is {abs(d_pts):.1f} pts below its "
                   f"trailing-{len(base)} average ({avg * 100:.1f}%)",
                   prior=avg, current=cur, kind="rate", change_str=_pts(d_pts / 100))]
    return [_a("baseline_above", "INFO", "OB", "Connect rate", basis,
               f"Connect rate {cur * 100:.1f}% is {d_pts:.1f} pts above its "
               f"trailing-{len(base)} average ({avg * 100:.1f}%)",
               prior=avg, current=cur, kind="rate", change_str=_pts(d_pts / 100))]


def _record_alert(rec, hist, cfg):
    if not rec.get("ob"):
        return []
    look = int(cfg.get("record_lookback", 6) or 0)
    series = _prior_rates(hist, rec["client"], rec.get("period", "daily"))
    if look:
        series = series[-look:]
    if len(series) < 2:
        return []
    cur = rec["ob"]["rate"]
    span = f"last {len(series)}" if look else "on record"
    if cur > max(series):
        return [_a("record_high", "INFO", "OB", "Connect rate", "record",
                   f"Best connect rate {span}: {cur * 100:.1f}% "
                   f"(prev best {max(series) * 100:.1f}%)",
                   prior=max(series), current=cur, kind="rate")]
    if cur < min(series):
        return [_a("record_low", "WARN", "OB", "Connect rate", "record",
                   f"Worst connect rate {span}: {cur * 100:.1f}% "
                   f"(prev low {min(series) * 100:.1f}%)",
                   prior=min(series), current=cur, kind="rate")]
    return []


def _window_drift_alert(blk, cfg):
    if not (blk and cfg.get("best_window_drift_alert") and blk.get("best_window_shift")):
        return []
    pw, cw = blk["best_window_shift"]
    return [_a("best_window_drift", "WARN", "OB", "Best window", blk.get("basis", "—"),
               f"Best calling window shifted {pw} → {cw} — recheck dialer scheduling",
               window=cw)]


def _target_alerts(rec, cfg, targets):
    out = []
    t = (targets or {}).get(rec["client"])
    if not t:
        return out
    ob, ib = rec.get("ob"), rec.get("ib")
    if ob and t.get("connect_rate") is not None and ob.get("rate") is not None:
        if ob["rate"] < t["connect_rate"]:
            out.append(_a("target_connect", "WARN", "OB", "Connect rate", "vs target",
                          f"Connect rate {ob['rate'] * 100:.1f}% below target "
                          f"{t['connect_rate'] * 100:.0f}%",
                          current=ob["rate"], kind="rate"))
    if ib and t.get("max_miss_rate") is not None and ib.get("miss_rate") is not None:
        if ib["miss_rate"] > t["max_miss_rate"]:
            out.append(_a("target_miss", "WARN", "IB", "Miss rate", "vs target",
                          f"Miss rate {ib['miss_rate'] * 100:.1f}% above target "
                          f"max {t['max_miss_rate'] * 100:.0f}%",
                          current=ib["miss_rate"], kind="rate"))
    if (ob and rec.get("period") == "monthly" and t.get("monthly_spend") is not None
            and ob.get("total_cost") is not None):
        tgt, spend = t["monthly_spend"], ob["total_cost"]
        if spend > tgt:
            out.append(_a("target_spend", "WARN", "OB", "Total spend", "vs budget",
                          f"Spend ${spend:,.0f} over monthly budget ${tgt:,.0f}",
                          current=spend, kind="money"))
        elif spend >= 0.9 * tgt:
            out.append(_a("target_spend_pace", "INFO", "OB", "Total spend", "vs budget",
                          f"Spend ${spend:,.0f} at {spend / tgt * 100:.0f}% of "
                          f"${tgt:,.0f} budget",
                          current=spend, kind="money"))
    return out


# --------------------------------------------------------------------------- #
# Data quality (works with or without history)                                 #
# --------------------------------------------------------------------------- #
def _data_quality(rec, cfg):
    if not cfg.get("data_quality_checks", True):
        return []
    out = []
    ob, ib = rec.get("ob"), rec.get("ib")
    if ob:
        rate = ob.get("rate")
        if rate is not None and (rate < 0 or rate > 1):
            out.append(_a("dq_bad_rate", "CRITICAL", "Data", "Connect rate", "validation",
                          f"Connect rate {rate * 100:.1f}% is out of range (0–100%) — check the file"))
        if (ob.get("dials") or 0) == 0 and (ob.get("total_cost") or 0) > 0:
            out.append(_a("dq_spend_no_dials", "CRITICAL", "Data", "Dials", "validation",
                          f"${ob['total_cost']:,.2f} spend recorded with zero dials — check the file"))
    if ib:
        mr = ib.get("miss_rate")
        if mr is not None and (mr < 0 or mr > 1):
            out.append(_a("dq_bad_miss", "CRITICAL", "Data", "Miss rate", "validation",
                          f"Miss rate {mr * 100:.1f}% is out of range — check the file"))
    return out


# --------------------------------------------------------------------------- #
# Public: per-client evaluation                                                #
# --------------------------------------------------------------------------- #
def _disposition_alerts(blk, basis, cfg):
    """Movement in the inbound agent-disposition mix (Col S 'An Agent Call
    Response'). Headline = connected (C-*) share drop; plus any single
    disposition whose share moved past the threshold."""
    out = []
    moves = (blk or {}).get("ib_disp") or []
    if not moves:
        return out
    # headline: connected (C-*) share of dispositioned calls dropping
    pc = sum(m["prior_share"] for m in moves if str(m["code"]).startswith("C-"))
    cc = sum(m["current_share"] for m in moves if str(m["code"]).startswith("C-"))
    drop = (pc - cc) * 100
    if drop >= cfg.get("connected_disp_drop_pts", 5.0):
        out.append(_a("ib_connected_disp_drop", "WARN", "IB",
                      "Connected dispositions", basis,
                      f"Connected (C-*) share of agent dispositions fell "
                      f"{drop:.1f} pts {basis} ({pc * 100:.1f}% → {cc * 100:.1f}%)",
                      prior=pc, current=cc, kind="rate", change_str=_pts(cc - pc)))
    # individual disposition shifts
    shift = cfg.get("disposition_shift_pts", 5.0) / 100.0
    cap = int(cfg.get("disposition_max_alerts", 6) or 0)
    n = 0
    for m in moves:
        if abs(m["share_delta"]) < shift:
            continue
        if cap and n >= cap:
            break
        n += 1
        if m["is_new"]:
            verb = "appeared"
        elif m["is_gone"]:
            verb = "disappeared"
        else:
            verb = "up" if m["share_delta"] > 0 else "down"
        out.append(_a(f"disp_shift::{m['code']}", "INFO", "IB",
                      f"Disp {m['code']}", basis,
                      f"Disposition {m['code']} {verb} {basis}: "
                      f"{m['prior_share'] * 100:.1f}% → {m['current_share'] * 100:.1f}% "
                      f"({m['prior_count']:,} → {m['current_count']:,} calls)",
                      prior=m["prior_share"], current=m["current_share"],
                      kind="rate", change_str=_pts(m["share_delta"])))
    return out


def _agent_alerts(blk, basis, cfg):
    """Per-agent connect-rate movement (Col K/L). Eligible agents must handle
    >= agent_min_calls in BOTH periods. Drop = WARN, rise = INFO."""
    out = []
    moves = (blk or {}).get("ib_agents") or []
    if not moves:
        return out
    move_thr = cfg.get("agent_connect_move_pts", 8.0) / 100.0
    min_calls = int(cfg.get("agent_min_calls", 20) or 0)
    cap = int(cfg.get("agent_max_alerts", 5) or 0)
    n = 0
    for m in moves:
        if m["rate_delta"] is None:
            continue
        if min_calls and (m["prior_handled"] < min_calls
                          or m["current_handled"] < min_calls):
            continue
        if abs(m["rate_delta"]) < move_thr:
            continue
        if cap and n >= cap:
            break
        n += 1
        d = m["rate_delta"] * 100
        sev = "WARN" if d < 0 else "INFO"
        direction = "up" if d > 0 else "down"
        out.append(_a(f"agent_move::{m['agent']}", sev, "IB",
                      f"Agent {m['agent']}", basis,
                      f"Agent {m['agent']} connect rate {direction} {abs(d):.1f} pts "
                      f"{basis} ({m['prior_rate'] * 100:.1f}% → "
                      f"{m['current_rate'] * 100:.1f}%, {m['current_handled']:,} calls)",
                      prior=m["prior_rate"], current=m["current_rate"],
                      kind="rate", change_str=_pts(m["rate_delta"])))
    return out


def evaluate(rec, comparisons, hist, cfg=None, targets=None):
    cfg = cfg if cfg is not None else _cfg()
    if not cfg.get("enabled", True):
        return []
    targets = targets if targets is not None else _targets()
    blk = comparisons[0] if comparisons else None
    basis = blk["basis"] if blk else (rec.get("period", "period").title())

    alerts = []
    alerts += _data_quality(rec, cfg)
    if blk:
        alerts += _delta_alerts(blk, basis, cfg)
        alerts += _window_drift_alert(blk, cfg)
        alerts += _disposition_alerts(blk, basis, cfg)
        alerts += _agent_alerts(blk, basis, cfg)
    alerts += _level_alerts(rec, cfg)
    alerts += _streak_alert(rec, hist, cfg)
    alerts += _baseline_alert(rec, hist, cfg)
    alerts += _record_alert(rec, hist, cfg)
    alerts += _target_alerts(rec, cfg, targets)
    return _sorted(alerts)


# --------------------------------------------------------------------------- #
# Public: cross-client (portfolio) evaluation                                  #
# --------------------------------------------------------------------------- #
def evaluate_portfolio(records, movement, hist, cfg=None):
    cfg = cfg if cfg is not None else _cfg()
    if not cfg.get("enabled", True):
        return []
    ob_recs = [r for r in records if r.get("ob")]
    if len(ob_recs) < 2:
        return []
    basis = {"weekly": "WoW", "monthly": "MoM"}.get(
        ob_recs[0].get("period", "daily"), "DoD")

    cur_dials = sum(r["ob"]["dials"] for r in ob_recs)
    cur_conn = sum(r["ob"]["connects"] for r in ob_recs)
    cur_blend = (cur_conn / cur_dials) if cur_dials else 0
    alerts = []

    # ---- portfolio blended connect drop vs prior period ----
    pd_dials = pd_conn = 0
    for r in ob_recs:
        period = r.get("period", "daily")
        recs = hist.get(r["client"], {}).get(period, [])
        p = trends._most_recent_before(recs, r.get("period_start"))
        if p and p.get("ob"):
            pd_dials += p["ob"]["dials"]
            pd_conn += p["ob"]["connects"]
    if pd_dials:
        prior_blend = pd_conn / pd_dials
        drop_pts = (prior_blend - cur_blend) * 100
        if drop_pts >= cfg.get("portfolio_blended_drop_pts", 2.0):
            alerts.append(_a("portfolio_blended_drop", "WARN", "Portfolio",
                             "Blended connect", basis,
                             f"Portfolio blended connect down {drop_pts:.1f} pts {basis} "
                             f"({prior_blend * 100:.1f}% → {cur_blend * 100:.1f}%)",
                             prior=prior_blend, current=cur_blend, kind="rate",
                             change_str=_pts(-drop_pts / 100)))

    # ---- clients below the blended average ----
    gap = cfg.get("below_portfolio_pts", 5.0)
    for r in sorted(ob_recs, key=lambda x: x["ob"]["rate"]):
        d = (cur_blend - r["ob"]["rate"]) * 100
        if d >= gap:
            alerts.append(_a("below_portfolio", "WARN", "Portfolio", "Connect rate",
                             "vs blended avg",
                             f"{r['client']} at {r['ob']['rate'] * 100:.1f}% is {d:.1f} pts "
                             f"below the {cur_blend * 100:.1f}% portfolio average",
                             prior=cur_blend, current=r["ob"]["rate"], kind="rate"))

    # ---- rank drops vs prior period ----
    spots = int(cfg.get("rank_drop_spots", 2) or 0)
    for c, m in (movement or {}).items():
        rc = m.get("rank_change")
        if rc is not None and spots and rc <= -spots:
            alerts.append(_a("rank_drop", "WARN", "Portfolio", "Rank", basis,
                             f"{c} dropped {abs(rc)} rank(s) {basis} "
                             f"(#{m.get('rank_prior')} → #{m.get('rank_now')})",
                             change_str=f"↓{abs(rc)}"))

    # ---- current worst performer ----
    worst = min(ob_recs, key=lambda x: x["ob"]["rate"])
    best = max(ob_recs, key=lambda x: x["ob"]["rate"])
    if worst["client"] != best["client"]:
        alerts.append(_a("worst_performer", "INFO", "Portfolio", "Connect rate",
                         "this period",
                         f"Lowest connect rate this period: {worst['client']} "
                         f"({worst['ob']['rate'] * 100:.1f}%)",
                         current=worst["ob"]["rate"], kind="rate"))
    return _sorted(alerts)


# --------------------------------------------------------------------------- #
# Slack digest formatting (plain text — Workflow webhook doesn't render mrkdwn)#
# --------------------------------------------------------------------------- #
def _min_rank(cfg):
    return SEV_RANK.get((cfg or _cfg()).get("slack_min_severity", "INFO"), 2)


def client_digest(client, period, alerts, primary_move=None, cfg=None):
    """Leadership change-summary posted to the report channel after the report."""
    cfg = cfg if cfg is not None else _cfg()
    basis = {"weekly": "Week-over-week", "monthly": "Month-over-month"}.get(
        period, "Day-over-day")
    keep = [a for a in alerts if SEV_RANK.get(a["severity"], 9) <= _min_rank(cfg)]
    lines = ["",
             f"🔔  {client} — Change Summary ({basis})",
             "━━━━━━━━━━━━━━━━━━━━"]
    if keep:
        for a in keep:
            lines.append(f"{SEV_EMOJI.get(a['severity'], '•')} {a['message']}")
        n_crit = sum(1 for a in keep if a["severity"] == "CRITICAL")
        n_warn = sum(1 for a in keep if a["severity"] == "WARN")
        lines.append("")
        lines.append(f"     {n_crit} critical · {n_warn} warning · {len(keep)} total flags")
    else:
        if primary_move:
            mv = primary_move
            lines.append(f"✅ No threshold alerts. Biggest move: connect "
                         f"{mv['arrow']} {mv['delta_str']} "
                         f"({mv['prior']:.1%} → {mv['current']:.1%}).")
        else:
            lines.append("✅ No threshold alerts this period.")
    return "\n".join(lines)


def portfolio_digest(period, alerts, cfg=None):
    """Portfolio-level alert block appended after the roll-up message."""
    cfg = cfg if cfg is not None else _cfg()
    keep = [a for a in alerts if SEV_RANK.get(a["severity"], 9) <= _min_rank(cfg)]
    if not keep:
        return ""
    basis = {"weekly": "Week-over-week", "monthly": "Month-over-month"}.get(
        period, "Day-over-day")
    lines = ["",
             f"🔔  PORTFOLIO ALERTS ({basis})",
             "━━━━━━━━━━━━━━━━━━━━"]
    for a in keep:
        lines.append(f"{SEV_EMOJI.get(a['severity'], '•')} {a['message']}")
    return "\n".join(lines)
