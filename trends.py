#!/usr/bin/env python3
"""
Trend layer for the TCN reports — Week-over-Week, Month-over-Month, Day-over-Day.

Each run that produces a report appends a compact metrics record to history.json.
From the second period of the same type onward we can compare the current period
to the right prior period and report raw deltas (absolute + %), with a direction
marker. No significance filtering — every move is shown as-is (per request).

History layout (JSON):
    { "<client>": { "weekly": [rec, rec, ...],
                    "monthly": [...],
                    "daily": [...] } }
Each rec is the client_metrics() dict + "period_start" / "period_end" (YYYY-MM-DD).

Comparison chosen by the CURRENT file's period:
    weekly  -> Week-over-week   (immediately prior weekly record)
    monthly -> Month-over-month (immediately prior monthly record)
    daily   -> Day-over-day     (immediately prior daily record)
               + vs same weekday last week (record dated current_start - 7 days)
"""

import json
import os
from datetime import date, datetime, timedelta

import config


# --------------------------------------------------------------------------- #
# History persistence                                                          #
# --------------------------------------------------------------------------- #
def _history_file():
    return getattr(config, "HISTORY_FILE", "history.json")


def load_history():
    path = _history_file()
    if os.path.exists(path):
        try:
            with open(path) as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def save_history(hist):
    with open(_history_file(), "w") as fh:
        json.dump(hist, fh, indent=2, default=str)


def append_record(hist, rec):
    """Add a record, replacing any existing one with the same period_start
    (so reprocessing the same period overwrites rather than duplicates)."""
    client = rec["client"]
    period = rec.get("period", "daily")
    bucket = hist.setdefault(client, {}).setdefault(period, [])
    ps = rec.get("period_start")
    bucket[:] = [r for r in bucket if r.get("period_start") != ps]
    bucket.append(rec)
    bucket.sort(key=lambda r: r.get("period_start") or "")
    return hist


# --------------------------------------------------------------------------- #
# Date helpers                                                                 #
# --------------------------------------------------------------------------- #
def _to_date(s):
    if not s:
        return None
    if isinstance(s, (date, datetime)):
        return s if isinstance(s, date) and not isinstance(s, datetime) else s.date()
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _most_recent_before(records, cur_start):
    cs = _to_date(cur_start)
    cands = [r for r in records if _to_date(r.get("period_start")) and
             (cs is None or _to_date(r["period_start"]) < cs)]
    if not cands:
        return None
    return max(cands, key=lambda r: _to_date(r["period_start"]))


def _exact_start(records, target):
    t = _to_date(target)
    for r in records:
        if _to_date(r.get("period_start")) == t:
            return r
    return None


def _minus_days(start, n):
    d = _to_date(start)
    return (d - timedelta(days=n)) if d else None


# --------------------------------------------------------------------------- #
# Delta computation                                                            #
# --------------------------------------------------------------------------- #
# (label, stream, key, fmt, kind, higher_better)
#   kind "rate"  -> delta shown in percentage POINTS
#   kind "count"/"money" -> delta shown in native units, %Δ relative
OB_METRICS = [
    ("Connect rate",     "ob", "rate",              "0.0%",        "rate",  True),
    ("Connects",         "ob", "connects",          "#,##0",       "count", True),
    ("Dials",            "ob", "dials",             "#,##0",       "count", None),
    ("Dials / connect",  "ob", "dials_per_connect", "0.0",         "count", False),
    ("$ / connect",      "ob", "cost_per_connect",  "$#,##0.000",  "money", False),
    ("Total spend",      "ob", "total_cost",        "$#,##0",      "money", None),
]
IB_METRICS = [
    ("Connect rate",     "ib", "rate",       "0.0%",  "rate",  True),
    ("Miss rate",        "ib", "miss_rate",  "0.0%",  "rate",  False),
    ("Missed",           "ib", "missed",     "#,##0", "count", False),
    ("Calls",            "ib", "calls",      "#,##0", "count", None),
]


def _arrow(delta, higher_better):
    if delta is None or abs(delta) < 1e-12:
        return "▬"
    up = delta > 0
    if higher_better is None:
        return "▲" if up else "▼"
    good = (up and higher_better) or (not up and not higher_better)
    return ("▲" if up else "▼")  # arrow = direction; color handles good/bad


def _is_good(delta, higher_better):
    if delta is None or abs(delta) < 1e-12 or higher_better is None:
        return None
    up = delta > 0
    return (up and higher_better) or (not up and not higher_better)


def _fmt_delta(delta, kind):
    if delta is None:
        return "—"
    sign = "+" if delta > 0 else ("" if delta == 0 else "−")
    a = abs(delta)
    if kind == "rate":
        return f"{sign}{a * 100:.1f} pts"
    if kind == "money":
        return f"{sign}${a:,.3f}"
    if a == int(a):
        return f"{sign}{int(a):,}"
    return f"{sign}{a:,.1f}"


def _fmt_pct(prior, cur):
    if prior in (None, 0) or cur is None:
        return "—"
    p = (cur - prior) / prior
    sign = "+" if p > 0 else ("" if p == 0 else "−")
    return f"{sign}{abs(p) * 100:.1f}%"


def _metric_rows(metrics, prior, cur):
    rows = []
    for label, stream, key, fmt, kind, hb in metrics:
        ps, cs = prior.get(stream), cur.get(stream)
        if not ps or not cs:
            continue
        pv, cv = ps.get(key), cs.get(key)
        if pv is None or cv is None:
            continue
        delta = cv - pv
        rows.append({
            "metric": label, "prior": pv, "current": cv,
            "delta": delta, "delta_str": _fmt_delta(delta, kind),
            "pct_str": _fmt_pct(pv, cv), "arrow": _arrow(delta, hb),
            "good": _is_good(delta, hb), "fmt": fmt, "kind": kind,
        })
    return rows


def _block(basis, prior, cur):
    block = {"basis": basis,
             "prior_label": prior.get("date_label", "prior"),
             "current_label": cur.get("date_label", "current"),
             "ob": _metric_rows(OB_METRICS, prior, cur),
             "ib": _metric_rows(IB_METRICS, prior, cur),
             "best_window_shift": None}
    pob, cob = prior.get("ob"), cur.get("ob")
    if pob and cob:
        pw, cw = pob.get("best_window"), cob.get("best_window")
        if pw and cw and pw != cw:
            block["best_window_shift"] = (pw, cw)
    return block


def compute_comparisons(rec, hist):
    """Return a list of comparison blocks for `rec` given prior history.
    Empty list on the first period of its type (nothing to compare to)."""
    period = rec.get("period", "daily")
    cur_start = rec.get("period_start")
    records = hist.get(rec["client"], {}).get(period, [])
    blocks = []
    if period == "weekly":
        p = _most_recent_before(records, cur_start)
        if p:
            blocks.append(_block("Week-over-week", p, rec))
    elif period == "monthly":
        p = _most_recent_before(records, cur_start)
        if p:
            blocks.append(_block("Month-over-month", p, rec))
    else:  # daily
        p = _most_recent_before(records, cur_start)
        if p:
            blocks.append(_block("Day-over-day", p, rec))
        sw = _exact_start(records, _minus_days(cur_start, 7))
        if sw:
            blocks.append(_block("vs same weekday last week", sw, rec))
    return blocks


def primary_connect_move(comparisons):
    """The OB connect-rate row of the first comparison block, for the
    Leadership View badge + Slack one-liner. None if unavailable."""
    if not comparisons:
        return None
    for r in comparisons[0]["ob"]:
        if r["metric"] == "Connect rate":
            return {"basis": comparisons[0]["basis"],
                    "prior": r["prior"], "current": r["current"],
                    "delta": r["delta"], "delta_str": r["delta_str"],
                    "arrow": r["arrow"], "good": r["good"]}
    return None


# --------------------------------------------------------------------------- #
# Cross-client movement (for the portfolio roll-up)                            #
# --------------------------------------------------------------------------- #
def rollup_movement(records, hist):
    """For each client record in this run, find its prior same-period connect
    rate and compute the pts delta + rank change vs the prior period.

    Returns {client: {"prior_rate", "delta_pts", "rank_now", "rank_prior",
                       "rank_change"}}."""
    ob_now = [r for r in records if r.get("ob")]
    if not ob_now:
        return {}

    # current ranking by connect rate (1 = best)
    cur_rank = {r["client"]: i + 1 for i, r in enumerate(
        sorted(ob_now, key=lambda r: r["ob"]["rate"], reverse=True))}

    # prior rate per client
    prior_rate = {}
    for r in ob_now:
        period = r.get("period", "daily")
        recs = hist.get(r["client"], {}).get(period, [])
        p = _most_recent_before(recs, r.get("period_start"))
        if p and p.get("ob"):
            prior_rate[r["client"]] = p["ob"]["rate"]

    # prior ranking (only clients that have a prior rate)
    prior_clients = sorted(prior_rate.items(), key=lambda kv: kv[1], reverse=True)
    prior_rank = {c: i + 1 for i, (c, _) in enumerate(prior_clients)}

    out = {}
    for r in ob_now:
        c = r["client"]
        pr = prior_rate.get(c)
        rk_now = cur_rank.get(c)
        rk_prior = prior_rank.get(c)
        out[c] = {
            "prior_rate": pr,
            "delta_pts": (r["ob"]["rate"] - pr) if pr is not None else None,
            "rank_now": rk_now,
            "rank_prior": rk_prior,
            "rank_change": (rk_prior - rk_now) if (rk_prior and rk_now) else None,
        }
    return out
