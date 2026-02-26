"""
Dutchmann WIP Reconciliation Tool  v2
Run with: streamlit run wip_reconciliation.py
Requires: pip install streamlit pandas openpyxl
"""

import re
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Dutchmann WIP Recon", page_icon="⇄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Mono',monospace;}
.block-container{padding-top:1.5rem;padding-bottom:2rem;}
.metric-card{background:#1c2333;border:1px solid #252d3d;border-radius:8px;padding:14px 18px;margin-bottom:8px;}
.metric-label{font-size:10px;color:#64748b;letter-spacing:.08em;font-weight:700;}
.metric-value{font-size:18px;font-weight:700;margin-top:4px;}
.metric-sub{font-size:11px;color:#64748b;margin-top:3px;}
div[data-testid="stDataFrame"]{border:1px solid #252d3d;border-radius:8px;}
.stButton>button{background:#3b82f6;color:white;border:none;font-family:'IBM Plex Mono',monospace;
  font-weight:700;letter-spacing:.05em;border-radius:8px;width:100%;}
.stButton>button:hover{background:#2563eb;}
.pm{background:#14532d;color:#22c55e;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.pd{background:#78350f;color:#f59e0b;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.pu{background:#7f1d1d;color:#ef4444;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
.pw{background:#1e1b4b;color:#6366f1;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt(n):
    if n is None or (isinstance(n, float) and pd.isna(n)): return "—"
    return f"R {abs(n):,.2f}{' CR' if n < 0 else ''}"

def norm(raw):
    return re.sub(r"[^\w\-/.]", "", str(raw).upper()).strip()

def keys_from(text):
    s = str(text).upper()
    hits = set()
    for p in [r"INA\d+", r"JBR[\s]?\d+(?:CN)?",
              r"\b(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\d{5}\b",
              r"INV[\s]?\d{3,}", r"RO\d+", r"[A-Z]{2}\d+\.\d+",
              r"\b\d{6,}\b", r"\b\d{4,5}\b"]:
        for m in re.findall(p, s):
            k = norm(m)
            if k and not k.isdigit() or (k.isdigit() and len(k) >= 4):
                hits.add(k)
    return list(hits)

def fval(v):
    try:
        f = float(v)
        return 0.0 if pd.isna(f) else f
    except: return 0.0


# ── WIP Parser ─────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_wip(fb):
    xl = pd.ExcelFile(BytesIO(fb), engine="openpyxl")
    rows = []
    cost_types = {"purchases","cb 3 payments","cb 2 payments","cb payments",
                  "jbr001 - jb racing","feb24019","purchase returns"}
    for sheet in ["Jobs Closed","WIP"]:
        if sheet not in xl.sheet_names: continue
        df = pd.read_excel(BytesIO(fb), sheet_name=sheet, header=None, engine="openpyxl")
        for _, row in df.iterrows():
            et = str(row.iloc[1] if len(row)>1 else "").strip().lower()
            # Handle rows where date leaked into EntryType column
            is_cost = any(t in et for t in cost_types)
            # Also catch rows where col1 is a date and col2 has "Purchases" equivalent
            # by checking if col4 has a numeric amount and col5 has a project
            if not is_cost:
                # Try: col0=date, col1=entrytype_or_ref, col2=ref, col3=supplier, col4=amount
                # Some donor rows have date in col0, date again in col1 (data shift)
                # We detect these if col4 is numeric and col5 references a project
                try:
                    amt_check = float(row.iloc[4])
                    proj_check = str(row.iloc[5] if len(row)>5 else "")
                    if "project" in proj_check.lower() and abs(amt_check) > 0:
                        is_cost = True
                except: pass
            if not is_cost: continue

            ref      = str(row.iloc[2] if len(row)>2 else "").strip()
            supplier = str(row.iloc[3] if len(row)>3 else "").strip()
            try: debit = float(row.iloc[4])
            except: continue
            project  = str(row.iloc[5] if len(row)>5 else "")
            if not project or project.lower() in ("nan","none",""): continue

            rows.append({"date": str(row.iloc[0])[:10], "ref": ref,
                         "ref_key": norm(ref), "supplier": supplier,
                         "amount": debit, "project": project, "sheet": sheet})
    return pd.DataFrame(rows)


# ── Tracker Parser ─────────────────────────────────────────────────────────────
# Column semantics (consistent across all sheets observed):
#   col 1 = Date
#   col 2 = Description
#   col 3 = Inv/Ref
#   col 4 = Labour cost
#   col 5 = Parts cost
#   col 6 = Misc/Other cost  (also used for Donor amounts)
#   col 8 = Income (progress payments) — SKIP these rows
#   col 9 = tick/x marker
#
# Rules:
#   1. Cost = sum(col4, col5, col6) for a row — but only if col8 is empty (not income)
#   2. Consecutive rows with no date but same/continuation inv ref = split of same invoice
#   3. Skip: estimates, progress payments (col8 has value), incoming payments

SKIP_RE = re.compile(
    r"progress\s*(payment|invoice)|incoming\s*payment|inward\s*payment|"
    r"estimate$|cafe\s*9\s*estimate|rs\s*estimate|rs\s*auto\s*estimate|"
    r"build\s*slot|remittance", re.IGNORECASE)

INCOME_DESCS = re.compile(
    r"progress|incoming|inward|remit|invoice\s*#\d|deposit.*dmann|dmann.*deposit|"
    r"payment.*stage|stage.*payment", re.IGNORECASE)

@st.cache_data(show_spinner=False)
def parse_tracker(fb):
    xl = pd.ExcelFile(BytesIO(fb), engine="openpyxl")
    projects = []

    for sname in xl.sheet_names:
        df = pd.read_excel(BytesIO(fb), sheet_name=sname, header=None, engine="openpyxl")
        data = df.values.tolist()

        cost_start = grand_total = income_total = -1
        vehicle = sname

        for i, row in enumerate(data):
            rs = "|".join(str(c) for c in row).upper()
            if "COST TRACKER" in rs and "LABOUR" in rs:
                cost_start = i + 3
            if "GRAND TOTAL" in rs:
                for c in row[1:]:
                    v = fval(c)
                    if v > 100000: grand_total = v; break
            if "TOTAL (LESS DONOR)" in rs and grand_total < 0:
                for c in row[1:]:
                    v = fval(c)
                    if v > 10000: grand_total = v; break
            if "TOTAL INCOME" in rs:
                for c in row[1:]:
                    v = fval(c)
                    if v > 100000: income_total = v; break
            if any(x in rs for x in ["/ CAFÉ 9","/ CAFE 9","/ RS AUTO","/ C9"]):
                parts = [str(c) for c in row if str(c).strip() not in ("nan","")]
                vehicle = " ".join(parts)[:80]

        if cost_start < 0: continue

        # Find end of cost section
        cost_end = min(cost_start + 90, len(data))
        for i in range(cost_start, cost_end):
            rs = "|".join(str(c) for c in data[i]).upper()
            if "TOTAL" in rs and i > cost_start + 5:
                vals = [fval(c) for c in data[i][3:8] if fval(c) != 0]
                if vals and max(abs(v) for v in vals) > 5000:
                    cost_end = i; break

        # Parse cost lines, merging split rows
        raw_lines = []
        for i in range(cost_start, cost_end):
            row = data[i]
            desc = str(row[2] if row[2] is not None else "").strip()
            inv  = str(row[3] if row[3] is not None else "").strip()
            date = str(row[1] if row[1] is not None else "")[:10]

            if not desc or desc.lower() in ("nan","none","date","description",""): continue
            if SKIP_RE.search(desc): continue

            # Check if this is an income row (col 8 has value, col 4/5/6 empty)
            income_val = fval(row[8] if len(row) > 8 else None)
            labour = fval(row[4] if len(row) > 4 else None)
            parts  = fval(row[5] if len(row) > 5 else None)
            misc   = fval(row[6] if len(row) > 6 else None)
            cost   = labour + parts + misc

            # Skip pure income rows
            if income_val != 0 and cost == 0: continue
            # Skip progress payments even if they have a small cost
            if INCOME_DESCS.search(desc) and income_val != 0: continue

            if cost == 0 and "credit" not in desc.lower() and "bakoven" not in desc.lower():
                continue

            raw_lines.append({"date": date, "desc": desc, "inv": inv,
                               "labour": labour, "parts": parts, "misc": misc,
                               "total": cost, "row_idx": i})

        # Merge consecutive split rows: if a row has no inv ref but same ref
        # as previous, or if description is a continuation (e.g. "Parts" after "Labour: strip")
        merged = []
        for rl in raw_lines:
            # If this row has no date AND no inv ref, try to merge with previous
            if rl["date"] in ("", "nan", "None") and not rl["inv"] and merged:
                prev = merged[-1]
                prev["labour"] += rl["labour"]
                prev["parts"]  += rl["parts"]
                prev["misc"]   += rl["misc"]
                prev["total"]  += rl["total"]
                prev["desc"]   += f" + {rl['desc']}" if rl["desc"] not in prev["desc"] else ""
            # If same inv ref as previous row → merge amounts
            elif rl["inv"] and merged and norm(rl["inv"]) == norm(merged[-1]["inv"]):
                prev = merged[-1]
                prev["labour"] += rl["labour"]
                prev["parts"]  += rl["parts"]
                prev["misc"]   += rl["misc"]
                prev["total"]  += rl["total"]
            else:
                merged.append(dict(rl))

        if merged:
            projects.append({"name": sname, "vehicle": vehicle,
                              "lines": merged, "grand_total": grand_total,
                              "income_total": income_total})
    return projects


# ── Matching engine ────────────────────────────────────────────────────────────
def reconcile(tracker_lines, wip_df):
    tl = [dict(l, matched=False, wip_refs=[], wip_amount=None, diff=None)
          for l in tracker_lines]
    wip = wip_df.copy(); wip["used"] = False

    def try_match_ref(r, multi=False):
        """Return (refs_list, amount, supplier) or None"""
        rn = norm(r)
        if rn in ("", "NAN"): return None
        # Exact key match
        mask = (~wip["used"]) & (wip["ref_key"] == rn)
        hits = wip[mask]
        if not hits.empty:
            amt = hits["amount"].sum()
            wip.loc[hits.index, "used"] = True
            return hits["ref"].tolist(), amt, hits["supplier"].iloc[0]
        # JBR fuzzy
        if "JBR" in rn:
            for k in wip[~wip["used"]]["ref_key"].unique():
                if "JBR" in k and (rn in k or k in rn):
                    mask2 = (~wip["used"]) & (wip["ref_key"] == k)
                    hits2 = wip[mask2]
                    if not hits2.empty:
                        amt = hits2["amount"].sum()
                        wip.loc[hits2.index, "used"] = True
                        return hits2["ref"].tolist(), amt, hits2["supplier"].iloc[0]
        return None

    # Pass 1: ref key match
    for line in tl:
        if line["total"] == 0: continue
        all_keys = keys_from(line["inv"] + " " + line["desc"])
        for k in all_keys:
            result = try_match_ref(k)
            if result:
                refs, amt, supp = result
                line["matched"]    = True
                line["wip_refs"]   = refs
                line["wip_amount"] = amt
                line["diff"]       = round(line["total"] - amt, 2)
                break

    # Pass 2: amount fallback (unique)
    for line in tl:
        if line["matched"] or line["total"] == 0: continue
        mask = (~wip["used"]) & (wip["amount"].sub(line["total"]).abs() < 1.0)
        hits = wip[mask]
        if len(hits) == 1:
            wip.loc[hits.index, "used"] = True
            line["matched"]    = True
            line["wip_refs"]   = hits["ref"].tolist()
            line["wip_amount"] = hits["amount"].iloc[0]
            line["diff"]       = round(line["total"] - hits["amount"].iloc[0], 2)

    return tl, wip[~wip["used"]].copy()


# ── Display helpers ────────────────────────────────────────────────────────────
def build_df(lines):
    rows = []
    for l in lines:
        if l["matched"] and (l["diff"] is None or abs(l["diff"]) < 0.05):
            status = "✓ Matched"
        elif l["matched"] and l["diff"] is not None and abs(l["diff"]) >= 0.05:
            status = "⚠ Diff"
        else:
            status = "✗ No match"
        diff_s = ""
        if l["diff"] is not None and abs(l["diff"]) >= 0.05:
            diff_s = f"+R {l['diff']:,.2f}" if l["diff"] > 0 else f"-R {abs(l['diff']):,.2f}"
        rows.append({"Date": l["date"], "Description": l["desc"][:70],
                     "Inv / Ref": l["inv"][:50], "Tracker Amt": l["total"],
                     "Status": status, "WIP Ref": ", ".join(l["wip_refs"]),
                     "WIP Amt": l["wip_amount"], "Diff": diff_s})
    return pd.DataFrame(rows)

def style_df(df):
    def rs(row):
        n = len(row); s = [""]*n; idx = {c:i for i,c in enumerate(row.index)}
        st = row.get("Status","")
        si = idx.get("Status")
        if si is not None:
            if "Matched" in st:   s[si] = "color:#22c55e;font-weight:700"
            elif "Diff" in st:    s[si] = "color:#f59e0b;font-weight:700"
            elif "No match" in st:s[si] = "color:#ef4444;font-weight:700"
        di = idx.get("Diff")
        if di is not None and row.get("Diff",""):
            s[di] = "color:#f59e0b;font-weight:700"
        ti = idx.get("Tracker Amt")
        if ti is not None:
            try:
                v = float(row["Tracker Amt"])
                s[ti] = "color:#ef4444;font-weight:600" if v<0 else "color:#e2e8f0;font-weight:600"
            except: pass
        wi = idx.get("WIP Amt")
        if wi is not None and row.get("WIP Amt") is not None:
            try:
                v = float(row["WIP Amt"])
                s[wi] = "color:#ef4444" if v<0 else "color:#6366f1"
            except: pass
        return s
    return df.style.apply(rs, axis=1).format(
        {"Tracker Amt": lambda x: fmt(x) if x is not None else "—",
         "WIP Amt":     lambda x: fmt(x) if x is not None else "—"})


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    st.markdown("## ⇄ Dutchmann WIP Reconciliation")
    st.markdown("---")

    if "result"      not in st.session_state: st.session_state.result      = None
    if "adjustments" not in st.session_state: st.session_state.adjustments = []

    with st.sidebar:
        st.markdown("**FILES**")
        wip_file     = st.file_uploader("WIP Ledger (.xlsx)", type=["xlsx","xls"])
        tracker_file = st.file_uploader("Project Tracker (.xlsx)", type=["xlsx","xls"])

        wip_df = tracker_data = None

        if wip_file:
            with st.spinner("Parsing WIP..."):
                wip_df = parse_wip(wip_file.read())
            st.success(f"✓ {len(wip_df)} cost lines")

        if tracker_file:
            with st.spinner("Parsing tracker..."):
                tracker_data = parse_tracker(tracker_file.read())
            st.success(f"✓ {len(tracker_data)} projects")

        sel_tracker = sel_wip = None

        if tracker_data:
            st.markdown("**TRACKER PROJECT**")
            sel_tracker = st.selectbox("", [p["name"] for p in tracker_data],
                                       label_visibility="collapsed")

        if wip_df is not None and not wip_df.empty:
            st.markdown("**WIP PROJECTS**")
            all_wip = sorted(wip_df["project"].unique())
            sel_wip = st.multiselect("", all_wip, default=all_wip,
                                     label_visibility="collapsed")

        st.markdown("---")
        run = st.button("▶ RUN RECONCILIATION",
                        disabled=(wip_df is None or tracker_data is None))

    if run and wip_df is not None and tracker_data and sel_tracker:
        proj = next((p for p in tracker_data if p["name"] == sel_tracker), None)
        if proj:
            fwip = wip_df[wip_df["project"].isin(sel_wip or [])]
            tl, unmatched = reconcile(proj["lines"], fwip)
            st.session_state.result = {"proj": proj, "tl": tl,
                                       "unmatched": unmatched, "name": sel_tracker}
            st.session_state.adjustments = []

    if not st.session_state.result:
        st.info("Upload both files, select a project and WIP projects, then click **Run Reconciliation**.")
        return

    res = st.session_state.result
    proj = res["proj"]; tl = res["tl"]; unmatched = res["unmatched"]

    tracker_total     = sum(l["total"] for l in tl)
    wip_match_total   = sum((l["wip_amount"] or 0) for l in tl if l["matched"])
    wip_unmatch_total = unmatched["amount"].sum() if not unmatched.empty else 0.0
    wip_total         = wip_match_total + wip_unmatch_total
    raw_gap           = wip_total - tracker_total

    adj_w = sum(a["amount"] for a in st.session_state.adjustments if a["side"]=="WIP")
    adj_t = sum(a["amount"] for a in st.session_state.adjustments if a["side"]=="Tracker")
    adj_gap = (wip_total + adj_w) - (tracker_total + adj_t)

    mc = sum(1 for l in tl if l["matched"] and (l["diff"] is None or abs(l["diff"])<0.05))
    md = sum(1 for l in tl if l["matched"] and l["diff"] is not None and abs(l["diff"])>=0.05)
    ut = sum(1 for l in tl if not l["matched"])
    uw = len(unmatched)

    st.markdown(f"### {proj['name']} — {proj['vehicle']}")

    c1,c2,c3,c4 = st.columns(4)
    for col, lbl, val, clr in [
        (c1,"TRACKER TOTAL",tracker_total,"#3b82f6"),
        (c2,"WIP TOTAL",wip_total,"#6366f1")]:
        with col:
            st.markdown(f'<div class="metric-card"><div class="metric-label">{lbl}</div>'
                        f'<div class="metric-value" style="color:{clr}">{fmt(val)}</div></div>',
                        unsafe_allow_html=True)

    gc = "#22c55e" if abs(raw_gap)<100 else "#f59e0b" if abs(raw_gap)<10000 else "#ef4444"
    gd = "WIP higher" if raw_gap>0.05 else "Tracker higher" if raw_gap<-0.05 else "✓ Balanced"
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-label">RAW GAP</div>'
                    f'<div class="metric-value" style="color:{gc}">{fmt(abs(raw_gap))}</div>'
                    f'<div class="metric-sub">{gd}</div></div>', unsafe_allow_html=True)

    ac = "#22c55e" if abs(adj_gap)<100 else "#f59e0b"
    ad = "WIP higher" if adj_gap>0.05 else "Tracker higher" if adj_gap<-0.05 else "✓ Balanced"
    with c4:
        st.markdown(f'<div class="metric-card"><div class="metric-label">ADJUSTED GAP</div>'
                    f'<div class="metric-value" style="color:{ac}">{fmt(abs(adj_gap))}</div>'
                    f'<div class="metric-sub">{ad}</div></div>', unsafe_allow_html=True)

    st.markdown(
        f'<span class="pm">✓ {mc} matched</span>&nbsp;&nbsp;'
        f'<span class="pd">⚠ {md} with diff</span>&nbsp;&nbsp;'
        f'<span class="pu">✗ {ut} tracker unmatched</span>&nbsp;&nbsp;'
        f'<span class="pw">◈ {uw} WIP only</span>',
        unsafe_allow_html=True)
    st.markdown("---")

    tab1,tab2,tab3 = st.tabs(["📋 All Lines","⚠️ Differences & Unmatched","◈ WIP Not in Tracker"])

    with tab1:
        st.dataframe(style_df(build_df(tl)), use_container_width=True, hide_index=True)

    with tab2:
        prob = [l for l in tl if not l["matched"] or (l["diff"] is not None and abs(l["diff"])>=0.05)]
        if prob:
            st.dataframe(style_df(build_df(prob)), use_container_width=True, hide_index=True)
            nd = sum(l["diff"] for l in prob if l["matched"] and l["diff"] is not None)
            nt = sum(l["total"] for l in prob if not l["matched"])
            st.markdown(f"**Matched with differences:** {fmt(nd)} net &nbsp;|&nbsp; **Tracker unmatched:** {fmt(nt)}")
        else:
            st.success("No differences or unmatched lines.")

    with tab3:
        if not unmatched.empty:
            disp = unmatched[["date","ref","supplier","amount","project","sheet"]].copy()
            disp.columns = ["Date","Ref","Supplier","Amount","Project","Source"]
            def sw(df):
                def rs(row):
                    s=[""]*len(row)
                    try:
                        v=float(row.get("Amount",0)); i=list(row.index).index("Amount")
                        s[i]="color:#ef4444;font-weight:600" if v<0 else "color:#6366f1;font-weight:600"
                    except: pass
                    return s
                return df.style.apply(rs,axis=1).format({"Amount":fmt})
            st.dataframe(sw(disp), use_container_width=True, hide_index=True)
            st.markdown(f"**Total WIP not in tracker:** {fmt(unmatched['amount'].sum())}")
        else:
            st.success("All WIP items matched.")

    # Manual adjustments
    st.markdown("---")
    st.markdown("### Manual Adjustments")
    ca,cb,cc,cd = st.columns([3,2,1.5,1])
    with ca: al = st.text_input("Description", placeholder="e.g. Missing labour INA22399")
    with cb: aa = st.number_input("Amount (negative=credit)", value=0.0, step=0.01, format="%.2f")
    with cc: asi = st.selectbox("Add to", ["WIP","Tracker"])
    with cd:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("＋ Add"):
            if al and aa != 0:
                st.session_state.adjustments.append({"label":al,"amount":aa,"side":asi})
                st.rerun()

    if st.session_state.adjustments:
        adj_df = pd.DataFrame([{"#":i,"Side":a["side"],"Description":a["label"],"Amount":a["amount"]}
                                for i,a in enumerate(st.session_state.adjustments)])
        ct,cd2 = st.columns([5,1])
        with ct:
            st.dataframe(adj_df.style.format({"Amount":fmt}).applymap(
                lambda v:"color:#6366f1" if v=="WIP" else "color:#f59e0b",subset=["Side"]),
                use_container_width=True, hide_index=True)
        with cd2:
            di = st.number_input("Remove #", min_value=0,
                                  max_value=max(0,len(st.session_state.adjustments)-1), step=1, value=0)
            if st.button("✕ Remove"):
                st.session_state.adjustments.pop(di); st.rerun()

        naw = sum(a["amount"] for a in st.session_state.adjustments if a["side"]=="WIP")
        nat = sum(a["amount"] for a in st.session_state.adjustments if a["side"]=="Tracker")
        ng  = (wip_total + naw) - (tracker_total + nat)
        nc  = "#22c55e" if abs(ng)<100 else "#f59e0b"
        nd2 = "WIP higher" if ng>0.05 else "Tracker higher" if ng<-0.05 else "✓ Balanced"
        st.markdown(f'<div class="metric-card" style="max-width:340px">'
                    f'<div class="metric-label">ADJUSTED GAP</div>'
                    f'<div class="metric-value" style="color:{nc}">{fmt(abs(ng))}</div>'
                    f'<div class="metric-sub">{nd2}</div></div>', unsafe_allow_html=True)

    # Export
    st.markdown("---")
    if st.button("⬇ Export to Excel"):
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            build_df(tl).to_excel(w, sheet_name="Tracker Lines", index=False)
            if not unmatched.empty:
                unmatched[["date","ref","supplier","amount","project"]].to_excel(
                    w, sheet_name="WIP Not in Tracker", index=False)
            if st.session_state.adjustments:
                pd.DataFrame(st.session_state.adjustments).to_excel(
                    w, sheet_name="Adjustments", index=False)
            pd.DataFrame([
                {"Item":"Tracker Total","Amount":tracker_total},
                {"Item":"WIP Total","Amount":wip_total},
                {"Item":"Raw Gap","Amount":raw_gap},
                {"Item":"Adj WIP","Amount":adj_w},
                {"Item":"Adj Tracker","Amount":adj_t},
                {"Item":"Adjusted Gap","Amount":adj_gap},
            ]).to_excel(w, sheet_name="Summary", index=False)
        st.download_button("📥 Download", data=out.getvalue(),
                           file_name=f"recon_{res['name']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
