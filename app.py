
import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import io
import os
import json
from datetime import datetime, date
import altair as alt

st.set_page_config(page_title="DC Performance Dashboard", page_icon="📦", layout="wide")

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 24px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 2rem; color: white; }
    .main-header p  { margin: 4px 0 0; opacity: .85; color: white; }

    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 18px 20px;
        border-left: 5px solid #2d6a9f;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        margin-bottom: 12px;
    }
    .metric-card h3 { margin: 0 0 4px; font-size: .85rem; color: #666; text-transform: uppercase; letter-spacing: .05em; }
    .metric-card p  { margin: 0; font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }

    .badge-trainee    { background:#ffeaa7; color:#d35400; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-starter    { background:#a8e6cf; color:#1e8449; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-competent  { background:#74b9ff; color:#1a5276; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-master     { background:#fd79a8; color:#6c1837; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }

    .user-detail-card {
        background: #f8fafc;
        border-radius: 10px;
        padding: 20px;
        border: 1px solid #e2e8f0;
        margin-bottom: 16px;
    }
    .section-header {
        background: #eef2f7;
        padding: 10px 16px;
        border-radius: 8px;
        font-weight: 700;
        color: #1e3a5f;
        margin: 16px 0 10px;
    }
    .hour-cell-high  { background:#d4edda; color:#155724; font-weight:700; border-radius:4px; padding:2px 6px; }
    .hour-cell-mid   { background:#fff3cd; color:#856404; font-weight:600; border-radius:4px; padding:2px 6px; }
    .hour-cell-low   { background:#f8d7da; color:#721c24; font-weight:600; border-radius:4px; padding:2px 6px; }
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_occupation_level(hire_date_raw):
    try:
        if pd.isna(hire_date_raw):
            return "Trainee"
        hire = pd.to_datetime(hire_date_raw)
        days = (pd.Timestamp.today() - hire).days
        months = days / 30.44
        if months >= 6:   return "Master"
        elif months >= 3: return "Competent"
        elif months >= 1: return "Starter"
        else:             return "Trainee"
    except:
        return "Trainee"

def badge_html(level):
    cls = level.lower()
    return f'<span class="badge-{cls}">{level}</span>'

def parse_man_hours(f):
    df_raw = pd.read_excel(f, header=None)
    day_row   = df_raw.iloc[8].tolist()
    col_names = df_raw.iloc[9].tolist()
    cols = []
    current_day = ""
    day_slots   = ["In", "Out", "Man Hrs", "Missing"]
    slot_idx    = 0
    for i, val in enumerate(col_names):
        dv = day_row[i]
        if isinstance(dv, str) and dv.strip():
            current_day = dv.strip()
            slot_idx = 0
        if isinstance(val, str) and val.strip():
            name = val.strip()
            if name in day_slots and current_day:
                cols.append(f"{current_day}_{name}")
                slot_idx += 1
            else:
                cols.append(name)
        else:
            cols.append(f"col_{i}")
    data = df_raw.iloc[10:].copy()
    data.columns = cols[:len(data.columns)]
    data = data[data["Site"].notna() & (data["Site"] != "Site")].copy()
    data = data[data["Name Surname"].notna()].copy()
    data = data[~data["Name Surname"].str.strip().eq("Total")].copy()
    data["Hire Date"] = pd.to_datetime(data["Hire Date"], errors="coerce")
    data["Occupation Level"] = data["Hire Date"].apply(get_occupation_level)
    def weeks_on_site(hd):
        try:
            if pd.isna(hd): return 0
            return max(0, round((pd.Timestamp.today() - pd.to_datetime(hd)).days / 7, 1))
        except: return 0
    data["Weeks On Site"] = data["Hire Date"].apply(weeks_on_site)
    period_cols = [c for c in data.columns if "Period Totals_Man Hrs" in c or c == "col_41"]
    if period_cols:
        data["Period Man Hrs"] = pd.to_numeric(data[period_cols[0]], errors="coerce")
    day_labels = ["Mon 25-May", "Tue 26-May", "Wed 27-May", "Thu 28-May",
                  "Fri 29-May", "Sat 30-May", "Sun 31-May"]
    for d in day_labels:
        col = f"{d}_Man Hrs"
        if col in data.columns:
            data[f"{d}_hrs"] = pd.to_numeric(data[col], errors="coerce").fillna(0)
        else:
            data[f"{d}_hrs"] = 0.0
    return data

def parse_rate_csv(f, report_type):
    df_raw = pd.read_csv(f, header=None)
    header_row = None
    date_val   = None
    target_val = None
    for i, row in df_raw.iterrows():
        r = row.tolist()
        if "User" in str(r[0]) or "create_user" in str(r[0]):
            header_row = i
        for cell in r:
            cell_str = str(cell)
            if "/" in cell_str and len(cell_str) == 10:
                try:
                    date_val = pd.to_datetime(cell)
                except: pass
            if "Target" in cell_str and "=" in cell_str:
                try:
                    target_val = cell_str
                except: pass
    if header_row is None:
        return pd.DataFrame(), date_val, target_val, []

    hour_row = df_raw.iloc[header_row - 1].tolist() if header_row > 0 else []
    base_cols  = ["username", "Hour", "Total"]
    hour_cols  = []
    raw_hours  = []
    for i in range(3, len(df_raw.columns)):
        hv = hour_row[i] if i < len(hour_row) else ""
        hv_str = str(hv).replace(".0","").strip()
        if hv_str not in ["", "nan"]:
            hour_cols.append(f"h_{hv_str}")
            raw_hours.append(hv_str)
        else:
            hour_cols.append(f"h_col{i}")
            raw_hours.append(f"col{i}")
    all_cols = base_cols + hour_cols

    data = df_raw.iloc[header_row + 1:].copy()
    data.columns = all_cols[:len(data.columns)]
    data = data[data["username"].notna()].copy()
    data = data[~data["username"].astype(str).str.lower().str.strip().eq("total")].copy()
    data["Total"] = pd.to_numeric(data["Total"], errors="coerce").fillna(0)
    for hc in hour_cols:
        if hc in data.columns:
            data[hc] = pd.to_numeric(data[hc], errors="coerce").fillna(0)
    data["report_type"] = report_type
    data["report_date"] = date_val if date_val else pd.Timestamp.today().normalize()
    data["_raw_hours"]  = str(raw_hours)
    return data, date_val, target_val, raw_hours

def load_zip_csv(uploaded_file, report_type):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
        fname = z.namelist()[0]
        with z.open(fname) as f:
            return parse_rate_csv(f, report_type)

def compute_uph(rate_data, man_hours_data):
    if rate_data.empty or man_hours_data.empty:
        return pd.DataFrame()
    uph_rows = []
    for _, row in rate_data.iterrows():
        username = str(row["username"]).strip().lower()
        total_units = row["Total"]
        rt = row["report_type"]
        rd = row["report_date"]
        mh_match = man_hours_data[
            man_hours_data["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username
        ]
        if mh_match.empty:
            continue
        mh = mh_match.iloc[0]
        day_hrs = None
        for dcol in man_hours_data.columns:
            if "_hrs" in dcol:
                v = mh.get(dcol, 0)
                if v and v > 0:
                    day_hrs = v
                    break
        if day_hrs is None or day_hrs == 0:
            day_hrs = 8.0
        uph_rows.append({
            "username":         username,
            "Name Surname":     mh.get("Name Surname", ""),
            "Company":          mh.get("Company Name", ""),
            "Occupation":       mh.get("Occupation", ""),
            "Department":       mh.get("Department", ""),
            "Hire Date":        mh.get("Hire Date", ""),
            "Weeks On Site":    mh.get("Weeks On Site", 0),
            "Occupation Level": mh.get("Occupation Level", "Trainee"),
            "Report Type":      rt,
            "Report Date":      rd,
            "Total Units":      total_units,
            "Hours Worked":     day_hrs,
            "UPH":              round(total_units / day_hrs, 1) if day_hrs > 0 else 0,
        })
    return pd.DataFrame(uph_rows)

# ─── Session state ────────────────────────────────────────────────────────────
if "man_hours_df"   not in st.session_state: st.session_state.man_hours_df   = None
if "rate_dfs"       not in st.session_state: st.session_state.rate_dfs       = []
if "rate_meta"      not in st.session_state: st.session_state.rate_meta      = {}
if "uph_df"         not in st.session_state: st.session_state.uph_df         = pd.DataFrame()
if "selected_user"  not in st.session_state: st.session_state.selected_user  = None

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📦 DC Performance Dashboard</h1>
  <p>Man Hours · Units Per Hour · Occupation Levels · Daily & Weekly Tracking</p>
</div>
""", unsafe_allow_html=True)

# ─── Sidebar – File Uploads ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Upload Daily Files")
    st.markdown("Upload reports every day to track performance.")

    mh_file   = st.file_uploader("⏱️ Auto Man Hours Report (.xlsx)", type=["xlsx"], key="mh")
    pick_file = st.file_uploader("🔵 Picking Rate Report (.zip)",    type=["zip"],  key="pick")
    pack_file = st.file_uploader("🟢 Packing Rate Report (.zip)",    type=["zip"],  key="pack")
    put_file  = st.file_uploader("🟡 Putaway Rate Report (.zip)",    type=["zip"],  key="put")
    recv_file = st.file_uploader("🟠 Receiving Rate Report (.zip)",  type=["zip"],  key="recv")

    process_btn = st.button("⚡ Process Files", use_container_width=True, type="primary")

    if process_btn:
        with st.spinner("Processing files..."):
            errors      = []
            rate_frames = []
            rate_meta   = {}

            if mh_file:
                try:
                    st.session_state.man_hours_df = parse_man_hours(mh_file)
                    st.success(f"✅ Man Hours: {len(st.session_state.man_hours_df)} employees")
                except Exception as e:
                    errors.append(f"Man Hours error: {e}")

            for file_obj, rtype in [(pick_file,"Picking"),(pack_file,"Packing"),(put_file,"Putaway"),(recv_file,"Receiving")]:
                if file_obj:
                    try:
                        df, dv, target, raw_hours = load_zip_csv(file_obj, rtype)
                        if not df.empty:
                            rate_frames.append(df)
                            rate_meta[rtype] = {"date": dv, "target": target, "hours": raw_hours}
                            st.success(f"✅ {rtype}: {len(df)} users | Date: {dv.date() if dv else 'N/A'}")
                    except Exception as e:
                        errors.append(f"{rtype} error: {e}")

            if rate_frames:
                st.session_state.rate_dfs  = rate_frames
                st.session_state.rate_meta = rate_meta
                all_rates = pd.concat(rate_frames, ignore_index=True)
                if st.session_state.man_hours_df is not None:
                    st.session_state.uph_df = compute_uph(all_rates, st.session_state.man_hours_df)

            for e in errors:
                st.error(e)

    st.markdown("---")
    st.markdown("**Occupation Levels**")
    st.markdown("🟡 **Trainee** – ≤ 2 weeks")
    st.markdown("🟢 **Starter** – > 1 month")
    st.markdown("🔵 **Competent** – ≥ 3 months")
    st.markdown("🩷 **Master** – ≥ 6 months")

# ─── Main Tabs ────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Overview",
    "👤 User Detail",
    "📅 Daily Hours",
    "🏆 Leaderboard",
    "📋 Day's Rate Detail"
])

# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">📊 Workforce Overview</div>', unsafe_allow_html=True)
    mh  = st.session_state.man_hours_df
    uph = st.session_state.uph_df
    if mh is None:
        st.info("👈 Upload files in the sidebar and click **Process Files** to get started.")
    else:
        c1, c2, c3, c4, c5 = st.columns(5)
        level_counts = mh["Occupation Level"].value_counts()
        with c1: st.markdown(f'<div class="metric-card"><h3>Total Employees</h3><p>{len(mh):,}</p></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><h3>🟡 Trainees</h3><p>{level_counts.get("Trainee",0)}</p></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-card"><h3>🟢 Starters</h3><p>{level_counts.get("Starter",0)}</p></div>', unsafe_allow_html=True)
        with c4: st.markdown(f'<div class="metric-card"><h3>🔵 Competent</h3><p>{level_counts.get("Competent",0)}</p></div>', unsafe_allow_html=True)
        with c5: st.markdown(f'<div class="metric-card"><h3>🩷 Masters</h3><p>{level_counts.get("Master",0)}</p></div>', unsafe_allow_html=True)

        st.markdown('<div class="section-header">🏢 Employees by Company</div>', unsafe_allow_html=True)
        comp_df = mh.groupby("Company Name").agg(Employees=("Name Surname","count"), Avg_Weeks=("Weeks On Site","mean")).reset_index().sort_values("Employees", ascending=False)
        comp_df["Avg_Weeks"] = comp_df["Avg_Weeks"].round(1)
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        if not uph.empty:
            st.markdown('<div class="section-header">⚡ Units Per Hour Summary by Report Type</div>', unsafe_allow_html=True)
            uph_sum = uph.groupby("Report Type").agg(Workers=("username","count"), Avg_UPH=("UPH","mean"), Max_UPH=("UPH","max"), Min_UPH=("UPH","min"), Total_Units=("Total Units","sum")).reset_index()
            for col in ["Avg_UPH","Max_UPH","Min_UPH"]:
                uph_sum[col] = uph_sum[col].round(1)
            st.dataframe(uph_sum, use_container_width=True, hide_index=True)

        st.markdown('<div class="section-header">📈 Occupation Level Distribution</div>', unsafe_allow_html=True)
        occ_df = mh["Occupation Level"].value_counts().reset_index()
        occ_df.columns = ["Level", "Count"]
        color_map = {"Trainee":"#ffeaa7","Starter":"#a8e6cf","Competent":"#74b9ff","Master":"#fd79a8"}
        chart = alt.Chart(occ_df).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6).encode(
            x=alt.X("Level:N", sort=["Trainee","Starter","Competent","Master"]),
            y=alt.Y("Count:Q"),
            color=alt.Color("Level:N", scale=alt.Scale(domain=list(color_map.keys()), range=list(color_map.values())), legend=None),
            tooltip=["Level","Count"]
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">👤 User Detail</div>', unsafe_allow_html=True)
    mh  = st.session_state.man_hours_df
    uph = st.session_state.uph_df
    if mh is None:
        st.info("👈 Upload files and process them first.")
    else:
        fc1, fc2, fc3 = st.columns(3)
        with fc1: dept_filter = st.selectbox("Department", ["All"] + sorted(mh["Department"].dropna().unique().tolist()))
        with fc2: comp_filter = st.selectbox("Company",    ["All"] + sorted(mh["Company Name"].dropna().unique().tolist()))
        with fc3: lvl_filter  = st.selectbox("Occupation Level", ["All","Trainee","Starter","Competent","Master"])

        display_mh = mh.copy()
        if dept_filter != "All": display_mh = display_mh[display_mh["Department"] == dept_filter]
        if comp_filter != "All": display_mh = display_mh[display_mh["Company Name"] == comp_filter]
        if lvl_filter  != "All": display_mh = display_mh[display_mh["Occupation Level"] == lvl_filter]

        if not uph.empty:
            uph_pivot = uph.groupby("username").agg(UPH=("UPH","mean"), Report_Type=("Report Type", lambda x: ", ".join(x.unique()))).reset_index()
            uph_pivot["UPH"] = uph_pivot["UPH"].round(1)
            display_mh = display_mh.merge(uph_pivot, left_on="Cust-Oracle Username", right_on="username", how="left")

        show_cols = ["Cust-Oracle Username","Name Surname","Company Name","Department","Occupation","Occupation Level","Weeks On Site","Hire Date"]
        if "UPH" in display_mh.columns: show_cols += ["UPH","Report_Type"]

        search = st.text_input("🔍 Search by name or username")
        if search:
            display_mh = display_mh[
                display_mh["Name Surname"].astype(str).str.lower().str.contains(search.lower()) |
                display_mh["Cust-Oracle Username"].astype(str).str.lower().str.contains(search.lower())
            ]

        st.markdown(f"**Showing {len(display_mh)} employees**")
        final_table = display_mh[show_cols].rename(columns={
            "Cust-Oracle Username":"Username","Name Surname":"Full Name","Company Name":"Company",
            "Occupation Level":"Level","Weeks On Site":"Weeks","Report_Type":"Role(s)"
        }).reset_index(drop=True)
        final_table["Hire Date"] = pd.to_datetime(final_table["Hire Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        st.dataframe(final_table, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown('<div class="section-header">🔎 Individual User Detail</div>', unsafe_allow_html=True)
        username_input = st.text_input("Enter Username (Oracle) to see full detail:")
        if username_input:
            user_mh = mh[mh["Cust-Oracle Username"].astype(str).str.lower().str.strip() == username_input.lower().strip()]
            if user_mh.empty:
                st.warning(f"No employee found with username **{username_input}**")
            else:
                row   = user_mh.iloc[0]
                level = row.get("Occupation Level", "Trainee")
                st.markdown(f"""
                <div class="user-detail-card">
                  <h2 style="margin:0 0 8px;">{row.get("Name Surname","")}</h2>
                  <p style="margin:2px 0;"><b>Username:</b> {row.get("Cust-Oracle Username","")}</p>
                  <p style="margin:2px 0;"><b>Company:</b> {row.get("Company Name","")}</p>
                  <p style="margin:2px 0;"><b>Department:</b> {row.get("Department","")}</p>
                  <p style="margin:2px 0;"><b>Occupation:</b> {row.get("Occupation","")}</p>
                  <p style="margin:2px 0;"><b>Hire Date:</b> {str(row.get("Hire Date",""))[:10]}</p>
                  <p style="margin:2px 0;"><b>Years of Service:</b> {row.get("Years of Service (Yrs/Mths)","")}</p>
                  <p style="margin:8px 0 0;"><b>Weeks on Site:</b> {row.get("Weeks On Site", 0)}&nbsp;&nbsp;
                  <b>Occupation Level:</b> {badge_html(level)}</p>
                </div>
                """, unsafe_allow_html=True)

                day_labels = ["Mon 25-May","Tue 26-May","Wed 27-May","Thu 28-May","Fri 29-May","Sat 30-May","Sun 31-May"]
                hrs_data = []
                total_hrs = 0
                for d in day_labels:
                    col = f"{d}_hrs"
                    v = row.get(col, 0) if col in row.index else 0
                    if pd.isna(v): v = 0
                    hrs_data.append({"Day": d, "Hours": v})
                    total_hrs += v
                hrs_df = pd.DataFrame(hrs_data)

                st.markdown('<div class="section-header">📅 Hours Worked This Week</div>', unsafe_allow_html=True)
                hr_chart = alt.Chart(hrs_df).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#2d6a9f").encode(
                    x="Day:N", y="Hours:Q", tooltip=["Day","Hours"]
                ).properties(height=220)
                st.altair_chart(hr_chart, use_container_width=True)
                c1, c2 = st.columns(2)
                c1.metric("Total Hours This Week", f"{total_hrs:.1f} hrs")
                c2.metric("Avg Daily Hours",       f"{total_hrs/7:.1f} hrs")

                if not uph.empty:
                    user_uph = uph[uph["username"].str.lower() == username_input.lower()]
                    if not user_uph.empty:
                        st.markdown('<div class="section-header">⚡ Units Per Hour Performance</div>', unsafe_allow_html=True)
                        uph_show = user_uph[["Report Type","Total Units","Hours Worked","UPH","Report Date"]].copy()
                        uph_show["Report Date"] = pd.to_datetime(uph_show["Report Date"]).dt.strftime("%Y-%m-%d")
                        st.dataframe(uph_show, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">📅 Daily & Weekly Hours Tracker</div>', unsafe_allow_html=True)
    mh = st.session_state.man_hours_df
    if mh is None:
        st.info("👈 Upload files and process them first.")
    else:
        day_labels = ["Mon 25-May","Tue 26-May","Wed 27-May","Thu 28-May","Fri 29-May","Sat 30-May","Sun 31-May"]
        dept_sel   = st.selectbox("Filter by Department", ["All"] + sorted(mh["Department"].dropna().unique().tolist()), key="daily_dept")
        disp = mh.copy()
        if dept_sel != "All": disp = disp[disp["Department"] == dept_sel]

        records = []
        for _, r in disp.iterrows():
            row_data = {"Username":r.get("Cust-Oracle Username",""),"Name":r.get("Name Surname",""),"Company":r.get("Company Name",""),"Level":r.get("Occupation Level","")}
            wk_total = 0
            for d in day_labels:
                col = f"{d}_hrs"
                v = r.get(col, 0) if col in r.index else 0
                if pd.isna(v): v = 0
                row_data[d] = round(v, 2)
                wk_total += v
            row_data["Weekly Total"] = round(wk_total, 2)
            records.append(row_data)

        hours_table = pd.DataFrame(records)
        hours_table = hours_table[hours_table["Weekly Total"] > 0].sort_values("Weekly Total", ascending=False)
        st.markdown(f"**{len(hours_table)} employees with recorded hours**")
        st.dataframe(hours_table, use_container_width=True, hide_index=True)

        st.markdown('<div class="section-header">📊 Daily Total Hours Across Team</div>', unsafe_allow_html=True)
        daily_totals = [{"Day": d, "Total Hours": hours_table[d].sum()} for d in day_labels if d in hours_table.columns]
        dt_df = pd.DataFrame(daily_totals)
        day_chart = alt.Chart(dt_df).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#27ae60").encode(
            x="Day:N", y="Total Hours:Q", tooltip=["Day","Total Hours"]
        ).properties(height=260)
        st.altair_chart(day_chart, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">🏆 UPH Leaderboard</div>', unsafe_allow_html=True)
    uph = st.session_state.uph_df
    if uph.empty:
        st.info("Upload rate reports and process them to see the leaderboard.")
    else:
        rtype_sel = st.selectbox("Report Type", ["All"] + sorted(uph["Report Type"].unique().tolist()))
        disp_uph  = uph.copy()
        if rtype_sel != "All": disp_uph = disp_uph[disp_uph["Report Type"] == rtype_sel]
        disp_uph = disp_uph.sort_values("UPH", ascending=False).reset_index(drop=True)
        disp_uph.index += 1
        show = disp_uph[["Name Surname","username","Company","Department","Occupation","Occupation Level","Weeks On Site","Report Type","Total Units","UPH"]].copy()
        show.columns = ["Name","Username","Company","Department","Occupation","Level","Weeks","Type","Units","UPH"]
        st.dataframe(show, use_container_width=True)

        top20 = disp_uph.head(20)
        lb_chart = alt.Chart(top20).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
            x=alt.X("UPH:Q"),
            y=alt.Y("Name Surname:N", sort="-x"),
            color=alt.Color("Occupation Level:N",
                scale=alt.Scale(domain=["Trainee","Starter","Competent","Master"],
                                range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8"])),
            tooltip=["Name Surname","UPH","Report Type","Company","Occupation Level"]
        ).properties(height=500, title="Top 20 Workers – Units Per Hour")
        st.altair_chart(lb_chart, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-header">📋 Day's Rate Detail – Hourly Breakdown Per User</div>', unsafe_allow_html=True)

    mh        = st.session_state.man_hours_df
    rate_dfs  = st.session_state.rate_dfs
    rate_meta = st.session_state.rate_meta

    if not rate_dfs:
        st.info("👈 Upload at least one rate report (Picking / Packing / Putaway / Receiving) and click **Process Files**.")
    else:
        # ── Report type selector ──
        available_types = list(rate_meta.keys())
        selected_type   = st.selectbox("📂 Select Report Type", available_types)

        rate_df = next((d for d in rate_dfs if d["report_type"].iloc[0] == selected_type), None)
        if rate_df is None:
            st.warning("No data found for this report type.")
        else:
            meta       = rate_meta.get(selected_type, {})
            report_date = meta.get("date")
            target_str  = meta.get("target", "")
            raw_hours   = meta.get("hours", [])

            # ── Info banner ──
            col_info1, col_info2, col_info3 = st.columns(3)
            with col_info1:
                st.markdown(f'<div class="metric-card"><h3>Report Type</h3><p style="font-size:1.2rem">{selected_type}</p></div>', unsafe_allow_html=True)
            with col_info2:
                date_str = report_date.strftime("%d %B %Y") if report_date else "N/A"
                st.markdown(f'<div class="metric-card"><h3>Report Date</h3><p style="font-size:1.2rem">{date_str}</p></div>', unsafe_allow_html=True)
            with col_info3:
                st.markdown(f'<div class="metric-card"><h3>Target</h3><p style="font-size:1.2rem">{target_str if target_str else "N/A"}</p></div>', unsafe_allow_html=True)

            # ── Hour columns in this report ──
            hour_cols = [c for c in rate_df.columns if c.startswith("h_") and not c.startswith("h_col")]

            # ── Filters ──
            f1, f2, f3 = st.columns(3)
            with f1:
                search_user = st.text_input("🔍 Search Username or Name", key="rate_search")
            with f2:
                dept_opts = ["All"]
                comp_opts = ["All"]
                if mh is not None:
                    dept_opts += sorted(mh["Department"].dropna().unique().tolist())
                    comp_opts += sorted(mh["Company Name"].dropna().unique().tolist())
                dept_f = st.selectbox("Department", dept_opts, key="rate_dept")
            with f3:
                comp_f = st.selectbox("Company", comp_opts, key="rate_comp")

            # ── Build merged table ──
            records = []
            for _, row in rate_df.iterrows():
                username   = str(row["username"]).strip()
                total_units = row["Total"]

                # Employee details from man hours
                name       = username
                company    = "—"
                dept       = "—"
                occupation = "—"
                hire_date  = "—"
                weeks      = "—"
                level      = "—"
                years_svc  = "—"

                if mh is not None:
                    mh_match = mh[mh["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username.lower()]
                    if not mh_match.empty:
                        mhr       = mh_match.iloc[0]
                        name      = mhr.get("Name Surname", username)
                        company   = mhr.get("Company Name", "—")
                        dept      = mhr.get("Department", "—")
                        occupation= mhr.get("Occupation", "—")
                        hire_date = str(mhr.get("Hire Date", ""))[:10]
                        weeks     = mhr.get("Weeks On Site", 0)
                        level     = mhr.get("Occupation Level", "Trainee")
                        years_svc = mhr.get("Years of Service (Yrs/Mths)", "—")

                rec = {
                    "Username":         username,
                    "Full Name":        name,
                    "Company":          company,
                    "Department":       dept,
                    "Occupation":       occupation,
                    "Hire Date":        hire_date,
                    "Weeks On Site":    weeks,
                    "Level":            level,
                    "Years of Service": years_svc,
                    "Total Units":      int(total_units),
                }

                # Add each hour's units
                for hc in hour_cols:
                    label = hc.replace("h_", "Hour ") + ":00"
                    rec[label] = int(row.get(hc, 0)) if not pd.isna(row.get(hc, 0)) else 0

                # Units per hour (total / active hours)
                active_hours = sum(1 for hc in hour_cols if row.get(hc, 0) > 0)
                rec["Active Hours"] = active_hours
                rec["Avg UPH"]      = round(total_units / active_hours, 1) if active_hours > 0 else 0

                records.append(rec)

            result_df = pd.DataFrame(records)

            # ── Apply filters ──
            if search_user:
                result_df = result_df[
                    result_df["Username"].str.lower().str.contains(search_user.lower()) |
                    result_df["Full Name"].str.lower().str.contains(search_user.lower())
                ]
            if dept_f != "All" and "Department" in result_df.columns:
                result_df = result_df[result_df["Department"] == dept_f]
            if comp_f != "All" and "Company" in result_df.columns:
                result_df = result_df[result_df["Company"] == comp_f]

            result_df = result_df.sort_values("Total Units", ascending=False).reset_index(drop=True)

            # ── Summary KPIs ──
            st.markdown('<div class="section-header">📊 Summary</div>', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("👷 Workers Active",  len(result_df))
            k2.metric("📦 Total Units",     f"{result_df['Total Units'].sum():,}")
            k3.metric("⚡ Avg UPH",          f"{result_df['Avg UPH'].mean():.1f}")
            k4.metric("🏆 Best UPH",         f"{result_df['Avg UPH'].max():.1f}")

            # ── Full table ──
            st.markdown(f'<div class="section-header">👥 All Workers – {selected_type} | {date_str}</div>', unsafe_allow_html=True)
            st.markdown(f"**{len(result_df)} workers found**")
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            # ── Individual user drill-down ──
            st.markdown("---")
            st.markdown('<div class="section-header">🔎 Click a Worker to See Hourly Breakdown</div>', unsafe_allow_html=True)

            all_usernames = result_df["Username"].tolist()
            selected_u    = st.selectbox("Select a worker:", ["— Select —"] + all_usernames, key="rate_user_select")

            if selected_u != "— Select —":
                user_row = result_df[result_df["Username"] == selected_u].iloc[0]

                # Profile card
                level_val = user_row.get("Level", "Trainee")
                st.markdown(f"""
                <div class="user-detail-card">
                  <h2 style="margin:0 0 8px;">{user_row["Full Name"]}</h2>
                  <div style="display:flex; flex-wrap:wrap; gap:24px; margin-top:8px;">
                    <div><b>Username:</b> {user_row["Username"]}</div>
                    <div><b>Company:</b> {user_row["Company"]}</div>
                    <div><b>Department:</b> {user_row["Department"]}</div>
                    <div><b>Occupation:</b> {user_row["Occupation"]}</div>
                    <div><b>Hire Date:</b> {user_row["Hire Date"]}</div>
                    <div><b>Years of Service:</b> {user_row["Years of Service"]}</div>
                    <div><b>Weeks on Site:</b> {user_row["Weeks On Site"]}</div>
                    <div><b>Level:</b> {badge_html(str(level_val))}</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                # Hourly breakdown
                hour_label_cols = [c for c in result_df.columns if c.startswith("Hour ")]
                if hour_label_cols:
                    st.markdown('<div class="section-header">⏰ Units Produced Per Hour</div>', unsafe_allow_html=True)

                    hour_data = []
                    for hc in hour_label_cols:
                        val = user_row.get(hc, 0)
                        hour_data.append({"Hour": hc, "Units": val})
                    hour_df = pd.DataFrame(hour_data)

                    # KPI row
                    uh1, uh2, uh3, uh4 = st.columns(4)
                    uh1.metric("📦 Total Units",   f"{int(user_row['Total Units']):,}")
                    uh2.metric("⏱️ Active Hours",  user_row["Active Hours"])
                    uh3.metric("⚡ Avg UPH",        user_row["Avg UPH"])
                    peak_h = hour_df.loc[hour_df["Units"].idxmax(), "Hour"] if not hour_df.empty else "—"
                    uh4.metric("🏆 Peak Hour",      peak_h)

                    # Bar chart
                    color_condition = alt.condition(
                        alt.datum.Units >= hour_df["Units"].quantile(0.66),
                        alt.value("#27ae60"),
                        alt.condition(
                            alt.datum.Units >= hour_df["Units"].quantile(0.33),
                            alt.value("#f39c12"),
                            alt.value("#e74c3c")
                        )
                    )
                    bar = alt.Chart(hour_df).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                        x=alt.X("Hour:N", sort=None, title="Hour of Day"),
                        y=alt.Y("Units:Q", title="Units Produced"),
                        color=color_condition,
                        tooltip=["Hour","Units"]
                    ).properties(height=300, title=f"Hourly Units — {user_row['Full Name']} ({selected_type})")
                    st.altair_chart(bar, use_container_width=True)

                    # Table view
                    hour_df_display = hour_df[hour_df["Units"] > 0].reset_index(drop=True)
                    st.dataframe(hour_df_display, use_container_width=True, hide_index=True)

            # ── UPH bar chart for all workers ──
            st.markdown('<div class="section-header">📈 Units Per Hour – All Workers Today</div>', unsafe_allow_html=True)
            chart_data = result_df[result_df["Avg UPH"] > 0].sort_values("Avg UPH", ascending=False).head(30)
            if not chart_data.empty:
                uph_bar = alt.Chart(chart_data).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X("Avg UPH:Q", title="Units Per Hour"),
                    y=alt.Y("Full Name:N", sort="-x", title="Worker"),
                    color=alt.Color("Level:N",
                        scale=alt.Scale(domain=["Trainee","Starter","Competent","Master","—"],
                                        range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8","#dfe6e9"])),
                    tooltip=["Full Name","Username","Company","Avg UPH","Total Units","Level","Active Hours"]
                ).properties(height=600, title=f"Top 30 Workers by UPH — {selected_type}")
                st.altair_chart(uph_bar, use_container_width=True)
