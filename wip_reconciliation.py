"""
Dutchmann WIP Reconciliation Tool
Run with: streamlit run wip_reconciliation.py
Requires: pip install streamlit pandas openpyxl
"""

import re
import streamlit as st
import pandas as pd
from io import BytesIO

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dutchmann WIP Recon",
    page_icon="⇄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'IBM Plex Mono', monospace; }
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .metric-card {
        background: #1c2333; border: 1px solid #252d3d;
        border-radius: 8px; padding: 14px 18px; margin-bottom: 8px;
    }
    .metric-label { font-size: 10px; color: #64748b; letter-spacing: 0.08em; font-weight: 700; }
    .metric-value { font-size: 18px; font-weight: 700; margin-top: 4px; }
    .metric-sub   { font-size: 11px; color: #64748b; margin-top: 3px; }
    .section-header {
        font-size: 11px; font-weight: 700; color: #64748b;
        letter-spacing: 0.1em; margin-bottom: 10px; margin-top: 20px;
    }
    div[data-testid="stDataFrame"] { border: 1px solid #252d3d; border-radius: 8px; }
    .stButton>button {
        background: #3b82f6; color: white; border: none;
        font-family: 'IBM Plex Mono', monospace; font-weight: 700;
        letter-spacing: 0.05em; border-radius: 8px; width: 100%;
    }
    .stButton>button:hover { background: #2563eb; }
    .pill-matched   { background:#14532d; color:#22c55e; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:700; }
    .pill-diff      { background:#78350f; color:#f59e0b; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:700; }
    .pill-unmatched { background:#7f1d1d; color:#ef4444; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:700; }
    .pill-wip       { background:#1e1b4b; color:#6366f1; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt_zar(n):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "—"
    sign = " CR" if n < 0 else ""
    return f"R {abs(n):,.2f}{sign}"


def normalise_ref(raw):
    if not raw and raw != 0:
        return ""
    return re.sub(r"[^\w\-/.]", "", str(raw).upper()).strip()


def extract_keys(text):
    if not text:
        return []
    s = str(text).upper()
    hits = set()
    for p in [
        r"INA\d+",
        r"INV[\s]?\d+",
        r"JBR[\s]?\d+(?:CN)?",
        r"[A-Z]{3}\d{5}",
        r"RO[_\s]?(\d+)",
        r"[A-Z]{2}\d+\.\d+",
        r"\d{6,}",
        r"\d{4,5}",
    ]:
        for m in re.findall(p, s):
            k = normalise_ref(m)
            if k:
                hits.add(k)
    return list(hits)


# ── Parsers ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_wip(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes))
    rows = []
    purchase_types = {
        "purchases", "cb 3 payments", "cb 2 payments", "cb payments",
        "jbr001 - jb racing", "feb24019",
    }
    for sheet in ["Jobs Closed", "WIP"]:
        if sheet not in xl.sheet_names:
            continue
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=None)
        for _, row in df.iterrows():
            entry_type = str(row.iloc[1] if len(row) > 1 else "").strip().lower()
            if not any(t in entry_type for t in purchase_types):
                continue
            ref      = str(row.iloc[2] if len(row) > 2 else "").strip()
            supplier = str(row.iloc[3] if len(row) > 3 else "").strip()
            try:
                debit = float(row.iloc[4])
            except (TypeError, ValueError):
                continue
            project = str(row.iloc[5] if len(row) > 5 else "")
            if not project or project.lower() in ("nan", "none", ""):
                continue
            rows.append({
                "date":    str(row.iloc[0])[:10],
                "ref":     ref,
                "ref_key": normalise_ref(ref),
                "supplier": supplier,
                "amount":  debit,
                "project": project,
                "sheet":   sheet,
            })
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def parse_tracker(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes))
    projects = []

    for sheet_name in xl.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        data = df.values.tolist()

        cost_start  = -1
        grand_total = None
        income_total = None
        vehicle     = sheet_name

        for i, row in enumerate(data):
            row_str = "|".join(str(c) for c in row).upper()

            if "COST TRACKER" in row_str and "LABOUR" in row_str:
                cost_start = i + 3

            if "GRAND TOTAL" in row_str:
                for c in row[1:]:
                    try:
                        v = float(c)
                        if v > 100000:
                            grand_total = v; break
                    except (TypeError, ValueError):
                        pass

            if "TOTAL (LESS DONOR)" in row_str:
                if grand_total is None:
                    for c in row[1:]:
                        try:
                            v = float(c)
                            if v > 10000:
                                grand_total = v; break
                        except (TypeError, ValueError):
                            pass

            if "TOTAL INCOME" in row_str:
                for c in row[1:]:
                    try:
                        v = float(c)
                        if v > 100000:
                            income_total = v; break
                    except (TypeError, ValueError):
                        pass

            if any(x in row_str for x in ["/ CAFÉ 9", "/ CAFE 9", "/ RS AUTO", "/ C9"]):
                parts = [str(c) for c in row if str(c).strip() not in ("nan", "")]
                vehicle = " ".join(parts)[:80]

        if cost_start < 0:
            continue

        cost_end = min(cost_start + 90, len(data))
        for i in range(cost_start, cost_end):
            row = data[i]
            row_str = "|".join(str(c) for c in row).upper()
            if "TOTAL" in row_str and i > cost_start + 5:
                vals = []
                for c in row[3:7]:
                    try: vals.append(float(c))
                    except (TypeError, ValueError): pass
                if vals and max(vals) > 1000:
                    cost_end = i; break

        skip_re = re.compile(
            r"progress\s*(invoice|payment)|inward\s*payment|incoming\s*payment|"
            r"estimate|cafe\s*9\s*estimate|rs\s*auto\s*estimate", re.IGNORECASE
        )

        lines = []
        for i in range(cost_start, cost_end):
            row = data[i]
            if len(row) < 4:
                continue
            desc = str(row[2] if row[2] is not None else "").strip()
            inv  = str(row[3] if row[3] is not None else "").strip()
            if not desc or desc.lower() in ("nan", "none", "description", ""):
                continue
            if skip_re.search(desc):
                amounts = []
                for c in row[4:7]:
                    try: amounts.append(float(c))
                    except (TypeError, ValueError): pass
                if not amounts or all(a == 0 for a in amounts):
                    continue
                if re.search(r"progress|inward|incoming", desc, re.I):
                    continue

            labour = parts_amt = misc = 0.0
            try: labour    = float(row[4]) if row[4] is not None else 0.0
            except (TypeError, ValueError): pass
            try: parts_amt = float(row[5]) if row[5] is not None else 0.0
            except (TypeError, ValueError): pass
            try: misc      = float(row[6]) if row[6] is not None else 0.0
            except (TypeError, ValueError): pass

            total = labour + parts_amt + misc
            if total == 0 and "credit" not in desc.lower():
                continue

            lines.append({
                "date":        str(row[1] if row[1] is not None else "")[:10],
                "description": desc,
                "inv_ref":     inv,
                "keys":        extract_keys(inv + " " + desc),
                "labour":      labour,
                "parts":       parts_amt,
                "misc":        misc,
                "total":       total,
            })

        if lines:
            projects.append({
                "name":         sheet_name,
                "vehicle":      vehicle,
                "lines":        lines,
                "grand_total":  grand_total,
                "income_total": income_total,
            })

    return projects


# ── Matching engine ────────────────────────────────────────────────────────────
def reconcile(tracker_lines, wip_df):
    t_lines = [dict(l, matched=False, wip_refs=[], wip_amount=None, diff=None)
               for l in tracker_lines]
    wip = wip_df.copy()
    wip["used"] = False

    # Pass 1: ref key match
    for tl in t_lines:
        if tl["total"] == 0:
            continue
        for key in tl["keys"]:
            mask = (~wip["used"]) & (wip["ref_key"] == key)
            hits = wip[mask]
            if not hits.empty:
                matched_amt = hits["amount"].sum()
                wip.loc[hits.index, "used"] = True
                tl["matched"]    = True
                tl["wip_refs"]   = hits["ref"].tolist()
                tl["wip_amount"] = matched_amt
                tl["diff"]       = round(tl["total"] - matched_amt, 2)
                break

    # Pass 2: amount fallback (unique)
    for tl in t_lines:
        if tl["matched"] or tl["total"] == 0:
            continue
        mask = (~wip["used"]) & (wip["amount"].sub(tl["total"]).abs() < 1.0)
        hits = wip[mask]
        if len(hits) == 1:
            wip.loc[hits.index, "used"] = True
            tl["matched"]    = True
            tl["wip_refs"]   = hits["ref"].tolist()
            tl["wip_amount"] = hits["amount"].iloc[0]
            tl["diff"]       = round(tl["total"] - hits["amount"].iloc[0], 2)

    return t_lines, wip[~wip["used"]].copy()


# ── Build display dataframe ────────────────────────────────────────────────────
def build_tracker_df(lines):
    rows = []
    for l in lines:
        if l["matched"] and (l["diff"] is None or abs(l["diff"]) < 0.05):
            status = "✓ Matched"
        elif l["matched"] and l["diff"] is not None and abs(l["diff"]) >= 0.05:
            status = "⚠ Diff"
        else:
            status = "✗ No match"

        diff_str = ""
        if l["diff"] is not None and abs(l["diff"]) >= 0.05:
            diff_str = f"+R {l['diff']:,.2f}" if l["diff"] > 0 else f"-R {abs(l['diff']):,.2f}"

        rows.append({
            "Date":        l["date"],
            "Description": l["description"][:70],
            "Inv / Ref":   l["inv_ref"][:50],
            "Tracker Amt": l["total"],
            "Status":      status,
            "WIP Ref":     ", ".join(l["wip_refs"]),
            "WIP Amt":     l["wip_amount"],
            "Diff":        diff_str,
        })
    return pd.DataFrame(rows)


def style_tracker_df(df):
    def row_style(row):
        n = len(row)
        styles = [""] * n
        idx = {col: i for i, col in enumerate(row.index)}

        s = row.get("Status", "")
        si = idx.get("Status")
        if si is not None:
            if s == "✓ Matched":
                styles[si] = "color:#22c55e;font-weight:700"
            elif s == "⚠ Diff":
                styles[si] = "color:#f59e0b;font-weight:700"
                di = idx.get("Diff")
                if di is not None:
                    styles[di] = "color:#f59e0b;font-weight:700"
            elif s == "✗ No match":
                styles[si] = "color:#ef4444;font-weight:700"

        ti = idx.get("Tracker Amt")
        if ti is not None:
            try:
                v = float(row["Tracker Amt"])
                styles[ti] = "color:#ef4444;font-weight:600" if v < 0 else "color:#e2e8f0;font-weight:600"
            except (TypeError, ValueError):
                pass

        wi = idx.get("WIP Amt")
        if wi is not None and row.get("WIP Amt") is not None:
            try:
                v = float(row["WIP Amt"])
                styles[wi] = "color:#ef4444" if v < 0 else "color:#6366f1"
            except (TypeError, ValueError):
                pass

        return styles

    return df.style.apply(row_style, axis=1).format({
        "Tracker Amt": lambda x: fmt_zar(x) if x is not None else "—",
        "WIP Amt":     lambda x: fmt_zar(x) if x is not None else "—",
    })


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    st.markdown("## ⇄ Dutchmann WIP Reconciliation")
    st.markdown("---")

    # Session state
    if "result"      not in st.session_state: st.session_state.result      = None
    if "adjustments" not in st.session_state: st.session_state.adjustments = []

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown('<div class="section-header">FILES</div>', unsafe_allow_html=True)
        wip_file     = st.file_uploader("WIP Ledger (.xlsx)", type=["xlsx","xls"], key="wip")
        tracker_file = st.file_uploader("Project Tracker (.xlsx)", type=["xlsx","xls"], key="tracker")

        wip_df       = None
        tracker_data = None

        if wip_file:
            with st.spinner("Parsing WIP ledger..."):
                wip_df = parse_wip(wip_file.read())
            st.success(f"✓ {len(wip_df)} cost lines loaded")

        if tracker_file:
            with st.spinner("Parsing tracker..."):
                tracker_data = parse_tracker(tracker_file.read())
            st.success(f"✓ {len(tracker_data)} projects found")

        selected_tracker      = None
        selected_wip_projects = []

        if tracker_data:
            st.markdown('<div class="section-header">TRACKER PROJECT</div>', unsafe_allow_html=True)
            selected_tracker = st.selectbox(
                "Select tracker sheet",
                [p["name"] for p in tracker_data],
                label_visibility="collapsed",
            )

        if wip_df is not None and not wip_df.empty:
            st.markdown('<div class="section-header">WIP PROJECTS TO INCLUDE</div>', unsafe_allow_html=True)
            all_wip = sorted(wip_df["project"].unique())
            selected_wip_projects = st.multiselect(
                "WIP projects", all_wip, default=all_wip,
                label_visibility="collapsed",
            )

        st.markdown("---")
        run = st.button("▶ RUN RECONCILIATION", disabled=(wip_df is None or tracker_data is None))

    # ── Run ────────────────────────────────────────────────────────────────────
    if run and wip_df is not None and tracker_data and selected_tracker:
        proj = next((p for p in tracker_data if p["name"] == selected_tracker), None)
        if proj:
            filtered_wip = wip_df[wip_df["project"].isin(selected_wip_projects)]
            t_lines, unmatched = reconcile(proj["lines"], filtered_wip)
            st.session_state.result = {
                "proj": proj, "t_lines": t_lines, "unmatched": unmatched,
                "tracker_name": selected_tracker,
            }
            st.session_state.adjustments = []

    if st.session_state.result is None:
        st.info("Upload both files, select a project and WIP projects, then click **Run Reconciliation**.")
        return

    res       = st.session_state.result
    proj      = res["proj"]
    t_lines   = res["t_lines"]
    unmatched = res["unmatched"]

    # ── Totals ─────────────────────────────────────────────────────────────────
    tracker_total       = sum(l["total"] for l in t_lines)
    wip_matched_total   = sum((l["wip_amount"] or 0) for l in t_lines if l["matched"])
    wip_unmatched_total = unmatched["amount"].sum() if not unmatched.empty else 0.0
    wip_total           = wip_matched_total + wip_unmatched_total
    raw_gap             = wip_total - tracker_total

    adj_wip     = sum(a["amount"] for a in st.session_state.adjustments if a["side"] == "WIP")
    adj_tracker = sum(a["amount"] for a in st.session_state.adjustments if a["side"] == "Tracker")
    adj_gap     = (wip_total + adj_wip) - (tracker_total + adj_tracker)

    matched_clean = sum(1 for l in t_lines if l["matched"] and (l["diff"] is None or abs(l["diff"]) < 0.05))
    matched_diff  = sum(1 for l in t_lines if l["matched"] and l["diff"] is not None and abs(l["diff"]) >= 0.05)
    unmatched_t   = sum(1 for l in t_lines if not l["matched"])
    unmatched_w   = len(unmatched)

    # ── Header ─────────────────────────────────────────────────────────────────
    st.markdown(f"### {proj['name']} — {proj['vehicle']}")

    c1, c2, c3, c4 = st.columns(4)
    for col, label, val, color in [
        (c1, "TRACKER TOTAL",   tracker_total, "#3b82f6"),
        (c2, "WIP TOTAL",       wip_total,     "#6366f1"),
    ]:
        with col:
            st.markdown(f"""<div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value" style="color:{color}">{fmt_zar(val)}</div>
            </div>""", unsafe_allow_html=True)

    gap_col = "#22c55e" if abs(raw_gap) < 100 else "#f59e0b" if abs(raw_gap) < 10000 else "#ef4444"
    gap_dir = "WIP higher" if raw_gap > 0.05 else "Tracker higher" if raw_gap < -0.05 else "✓ Balanced"
    with c3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">RAW GAP</div>
            <div class="metric-value" style="color:{gap_col}">{fmt_zar(abs(raw_gap))}</div>
            <div class="metric-sub">{gap_dir}</div>
        </div>""", unsafe_allow_html=True)

    adj_col = "#22c55e" if abs(adj_gap) < 100 else "#f59e0b"
    adj_dir = "WIP higher" if adj_gap > 0.05 else "Tracker higher" if adj_gap < -0.05 else "✓ Balanced"
    with c4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">ADJUSTED GAP</div>
            <div class="metric-value" style="color:{adj_col}">{fmt_zar(abs(adj_gap))}</div>
            <div class="metric-sub">{adj_dir}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown(
        f'<span class="pill-matched">✓ {matched_clean} matched</span>&nbsp;&nbsp;'
        f'<span class="pill-diff">⚠ {matched_diff} with diff</span>&nbsp;&nbsp;'
        f'<span class="pill-unmatched">✗ {unmatched_t} tracker unmatched</span>&nbsp;&nbsp;'
        f'<span class="pill-wip">◈ {unmatched_w} WIP only</span>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Tables ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["📋 All Lines", "⚠️ Differences & Unmatched", "◈ WIP Not in Tracker"])

    with tab1:
        st.dataframe(style_tracker_df(build_tracker_df(t_lines)), use_container_width=True, hide_index=True)

    with tab2:
        problem = [l for l in t_lines if not l["matched"] or (l["diff"] is not None and abs(l["diff"]) >= 0.05)]
        if problem:
            st.dataframe(style_tracker_df(build_tracker_df(problem)), use_container_width=True, hide_index=True)
            net_diff = sum(l["diff"] for l in problem if l["matched"] and l["diff"] is not None)
            unm_total = sum(l["total"] for l in problem if not l["matched"])
            st.markdown(f"**Matched with differences:** {fmt_zar(net_diff)} net &nbsp;|&nbsp; **Tracker lines with no WIP match:** {fmt_zar(unm_total)}")
        else:
            st.success("No differences or unmatched lines.")

    with tab3:
        if not unmatched.empty:
            disp = unmatched[["date","ref","supplier","amount","project","sheet"]].copy()
            disp.columns = ["Date","Ref","Supplier","Amount","Project","Source"]

            def style_wip_df(df):
                def rs(row):
                    styles = [""] * len(row)
                    try:
                        v = float(row.get("Amount", 0))
                        i = list(row.index).index("Amount")
                        styles[i] = "color:#ef4444;font-weight:600" if v < 0 else "color:#6366f1;font-weight:600"
                    except (ValueError, TypeError):
                        pass
                    return styles
                return df.style.apply(rs, axis=1).format({"Amount": fmt_zar})

            st.dataframe(style_wip_df(disp), use_container_width=True, hide_index=True)
            st.markdown(f"**Total WIP not in tracker:** {fmt_zar(unmatched['amount'].sum())}")
        else:
            st.success("All WIP items matched.")

    # ── Manual adjustments ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Manual Adjustments")
    st.caption("Add missing items to bridge the gap — e.g. a labour invoice not posted to WIP, or a tracker cost missing from the ledger.")

    ca, cb, cc, cd = st.columns([3, 2, 1.5, 1])
    with ca: adj_label  = st.text_input("Description", placeholder="e.g. Missing labour INA22214", key="adj_label")
    with cb: adj_amount = st.number_input("Amount (negative = credit)", value=0.0, step=0.01, format="%.2f", key="adj_amount")
    with cc: adj_side   = st.selectbox("Add to", ["WIP", "Tracker"], key="adj_side")
    with cd:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("＋ Add"):
            if adj_label and adj_amount != 0:
                st.session_state.adjustments.append({"label": adj_label, "amount": adj_amount, "side": adj_side})
                st.rerun()

    if st.session_state.adjustments:
        adj_rows = [{"#": i, "Side": a["side"], "Description": a["label"], "Amount": a["amount"]}
                    for i, a in enumerate(st.session_state.adjustments)]
        adj_df = pd.DataFrame(adj_rows)

        ct, cd2 = st.columns([5, 1])
        with ct:
            st.dataframe(
                adj_df.style.format({"Amount": fmt_zar}).applymap(
                    lambda v: "color:#6366f1" if v == "WIP" else "color:#f59e0b", subset=["Side"]
                ),
                use_container_width=True, hide_index=True,
            )
        with cd2:
            del_idx = st.number_input("Remove row #", min_value=0,
                                       max_value=max(0, len(st.session_state.adjustments) - 1),
                                       step=1, value=0)
            if st.button("✕ Remove"):
                st.session_state.adjustments.pop(del_idx)
                st.rerun()

        new_adj_wip     = sum(a["amount"] for a in st.session_state.adjustments if a["side"] == "WIP")
        new_adj_tracker = sum(a["amount"] for a in st.session_state.adjustments if a["side"] == "Tracker")
        new_gap = (wip_total + new_adj_wip) - (tracker_total + new_adj_tracker)
        ng_col = "#22c55e" if abs(new_gap) < 100 else "#f59e0b"
        ng_dir = "WIP higher" if new_gap > 0.05 else "Tracker higher" if new_gap < -0.05 else "✓ Balanced"
        st.markdown(f"""<div class="metric-card" style="max-width:340px">
            <div class="metric-label">ADJUSTED GAP</div>
            <div class="metric-value" style="color:{ng_col}">{fmt_zar(abs(new_gap))}</div>
            <div class="metric-sub">{ng_dir}</div>
        </div>""", unsafe_allow_html=True)

    # ── Export ──────────────────────────────────────────────────────────────────
    st.markdown("---")
    if st.button("⬇ Export to Excel"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            build_tracker_df(t_lines).to_excel(writer, sheet_name="Tracker Lines", index=False)
            if not unmatched.empty:
                unmatched[["date","ref","supplier","amount","project"]].to_excel(
                    writer, sheet_name="WIP Not in Tracker", index=False)
            if st.session_state.adjustments:
                pd.DataFrame(st.session_state.adjustments).to_excel(
                    writer, sheet_name="Adjustments", index=False)
            pd.DataFrame([
                {"Item": "Tracker Total", "Amount": tracker_total},
                {"Item": "WIP Total",     "Amount": wip_total},
                {"Item": "Raw Gap",       "Amount": raw_gap},
                {"Item": "Adj WIP",       "Amount": adj_wip},
                {"Item": "Adj Tracker",   "Amount": adj_tracker},
                {"Item": "Adjusted Gap",  "Amount": adj_gap},
            ]).to_excel(writer, sheet_name="Summary", index=False)
        st.download_button(
            "📥 Download",
            data=output.getvalue(),
            file_name=f"recon_{res['tracker_name']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
