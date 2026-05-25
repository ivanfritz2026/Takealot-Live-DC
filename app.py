
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
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_occupation_level(hire_date_raw):
    """Return badge label based on time-since-hire."""
    try:
        if pd.isna(hire_date_raw):
            return "Trainee"
        hire = pd.to_datetime(hire_date_raw)
        days = (pd.Timestamp.today() - hire).days
        weeks = days / 7
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
    """Parse auto man-hours Excel report."""
    df_raw = pd.read_excel(f, header=None)

    # Row 9 = column names, row 8 = day headers
    day_row   = df_raw.iloc[8].tolist()
    col_names = df_raw.iloc[9].tolist()

    # Build column name list
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

    # Occupation level
    data["Hire Date"] = pd.to_datetime(data["Hire Date"], errors="coerce")
    data["Occupation Level"] = data["Hire Date"].apply(get_occupation_level)

    # Weeks on site
    def weeks_on_site(hd):
        try:
            if pd.isna(hd): return 0
            return max(0, round((pd.Timestamp.today() - pd.to_datetime(hd)).days / 7, 1))
        except: return 0

    data["Weeks On Site"] = data["Hire Date"].apply(weeks_on_site)

    # Period total man hours (last columns)
    period_cols = [c for c in data.columns if "Period Totals_Man Hrs" in c or c == "col_41"]
    if period_cols:
        data["Period Man Hrs"] = pd.to_numeric(data[period_cols[0]], errors="coerce")

    # Collect daily man hours
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
    """Parse picking/packing/putaway/receiving CSV from zip."""
    df_raw = pd.read_csv(f, header=None)

    # Find the header row (contains "User" or "create_user")
    header_row = None
    date_val   = None
    for i, row in df_raw.iterrows():
        r = row.tolist()
        if "User" in str(r[0]) or "create_user" in str(r[0]):
            header_row = i
        # grab date from row above
        for cell in r:
            if isinstance(cell, str) and "/" in cell and len(cell) == 10:
                try:
                    date_val = pd.to_datetime(cell)
                except: pass

    if header_row is None:
        return pd.DataFrame(), date_val

    # Hour columns from the row above header
    hour_row = df_raw.iloc[header_row - 1].tolist() if header_row > 0 else []

    # Build columns
    base_cols  = ["username", "Hour", "Total"]
    hour_cols  = []
    for i in range(3, len(df_raw.columns)):
        hv = hour_row[i] if i < len(hour_row) else ""
        hour_cols.append(f"h_{hv}" if str(hv) not in ["", "nan"] else f"h_col{i}")
    all_cols = base_cols + hour_cols

    data = df_raw.iloc[header_row + 1:].copy()
    data.columns = all_cols[:len(data.columns)]

    # Drop total row
    data = data[data["username"].notna()].copy()
    data = data[~data["username"].astype(str).str.lower().str.strip().eq("total")].copy()
    data["Total"] = pd.to_numeric(data["Total"], errors="coerce").fillna(0)
    data["report_type"] = report_type
    data["report_date"] = date_val if date_val else pd.Timestamp.today().normalize()
    return data, date_val

def load_zip_csv(uploaded_file, report_type):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
        fname = z.namelist()[0]
        with z.open(fname) as f:
            return parse_rate_csv(f, report_type)

def compute_uph(rate_data, man_hours_data):
    """Merge rate totals with man hours to get units per hour."""
    if rate_data.empty or man_hours_data.empty:
        return pd.DataFrame()

    uph_rows = []
    for _, row in rate_data.iterrows():
        username = str(row["username"]).strip().lower()
        total_units = row["Total"]
        rt = row["report_type"]
        rd = row["report_date"]

        # Match in man hours
        mh_match = man_hours_data[
            man_hours_data["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username
        ]
        if mh_match.empty:
            continue
        mh = mh_match.iloc[0]

        # Try to find man hrs for the day
        day_hrs = None
        for dcol in man_hours_data.columns:
            if "_hrs" in dcol:
                v = mh.get(dcol, 0)
                if v and v > 0:
                    day_hrs = v
                    break
        if day_hrs is None or day_hrs == 0:
            day_hrs = 8.0  # fallback

        # Convert HH:MM to float hours if needed
        def to_float_hours(v):
            try:
                if isinstance(v, str) and ":" in v:
                    parts = v.split(":")
                    return int(parts[0]) + int(parts[1]) / 60
                return float(v) if v else 0
            except:
                return 0

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
if "man_hours_df"  not in st.session_state: st.session_state.man_hours_df  = None
if "rate_dfs"      not in st.session_state: st.session_state.rate_dfs      = []
if "uph_df"        not in st.session_state: st.session_state.uph_df        = pd.DataFrame()
if "selected_user" not in st.session_state: st.session_state.selected_user = None

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

    mh_file = st.file_uploader("⏱️ Auto Man Hours Report (.xlsx)", type=["xlsx"], key="mh")
    pick_file = st.file_uploader("🔵 Picking Rate Report (.zip)", type=["zip"], key="pick")
    pack_file = st.file_uploader("🟢 Packing Rate Report (.zip)", type=["zip"], key="pack")
    put_file  = st.file_uploader("🟡 Putaway Rate Report (.zip)", type=["zip"], key="put")
    recv_file = st.file_uploader("🟠 Receiving Rate Report (.zip)", type=["zip"], key="recv")

    process_btn = st.button("⚡ Process Files", use_container_width=True, type="primary")

    if process_btn:
        with st.spinner("Processing files..."):
            errors = []
            rate_frames = []

            # Man Hours
            if mh_file:
                try:
                    st.session_state.man_hours_df = parse_man_hours(mh_file)
                    st.success(f"✅ Man Hours: {len(st.session_state.man_hours_df)} employees")
                except Exception as e:
                    errors.append(f"Man Hours error: {e}")

            # Rate reports
            for file_obj, rtype in [(pick_file,"Picking"),(pack_file,"Packing"),(put_file,"Putaway"),(recv_file,"Receiving")]:
                if file_obj:
                    try:
                        df, dv = load_zip_csv(file_obj, rtype)
                        if not df.empty:
                            rate_frames.append(df)
                            st.success(f"✅ {rtype}: {len(df)} users | Date: {dv.date() if dv else 'N/A'}")
                    except Exception as e:
                        errors.append(f"{rtype} error: {e}")

            if rate_frames:
                st.session_state.rate_dfs = rate_frames
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
tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "👤 User Detail", "📅 Daily Hours", "🏆 Leaderboard"])

# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">📊 Workforce Overview</div>', unsafe_allow_html=True)

    mh = st.session_state.man_hours_df
    uph = st.session_state.uph_df

    if mh is None:
        st.info("👈 Upload files in the sidebar and click **Process Files** to get started.")
    else:
        # KPI row
        c1, c2, c3, c4, c5 = st.columns(5)
        level_counts = mh["Occupation Level"].value_counts()
        with c1:
            st.markdown(f'<div class="metric-card"><h3>Total Employees</h3><p>{len(mh):,}</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-card"><h3>🟡 Trainees</h3><p>{level_counts.get("Trainee",0)}</p></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card"><h3>🟢 Starters</h3><p>{level_counts.get("Starter",0)}</p></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="metric-card"><h3>🔵 Competent</h3><p>{level_counts.get("Competent",0)}</p></div>', unsafe_allow_html=True)
        with c5:
            st.markdown(f'<div class="metric-card"><h3>🩷 Masters</h3><p>{level_counts.get("Master",0)}</p></div>', unsafe_allow_html=True)

        # Company breakdown
        st.markdown('<div class="section-header">🏢 Employees by Company</div>', unsafe_allow_html=True)
        comp_df = mh.groupby("Company Name").agg(
            Employees=("Name Surname","count"),
            Avg_Weeks=("Weeks On Site","mean")
        ).reset_index().sort_values("Employees", ascending=False)
        comp_df["Avg_Weeks"] = comp_df["Avg_Weeks"].round(1)
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        # UPH summary if available
        if not uph.empty:
            st.markdown('<div class="section-header">⚡ Units Per Hour Summary by Report Type</div>', unsafe_allow_html=True)
            uph_sum = uph.groupby("Report Type").agg(
                Workers=("username","count"),
                Avg_UPH=("UPH","mean"),
                Max_UPH=("UPH","max"),
                Min_UPH=("UPH","min"),
                Total_Units=("Total Units","sum"),
            ).reset_index()
            uph_sum["Avg_UPH"] = uph_sum["Avg_UPH"].round(1)
            uph_sum["Max_UPH"] = uph_sum["Max_UPH"].round(1)
            uph_sum["Min_UPH"] = uph_sum["Min_UPH"].round(1)
            st.dataframe(uph_sum, use_container_width=True, hide_index=True)

        # Occupation distribution chart
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
    st.markdown('<div class="section-header">👤 User Detail – Click a Username to Drill Down</div>', unsafe_allow_html=True)

    mh = st.session_state.man_hours_df
    uph = st.session_state.uph_df

    if mh is None:
        st.info("👈 Upload files and process them first.")
    else:
        # Filter controls
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            dept_filter = st.selectbox("Department", ["All"] + sorted(mh["Department"].dropna().unique().tolist()))
        with fc2:
            comp_filter = st.selectbox("Company", ["All"] + sorted(mh["Company Name"].dropna().unique().tolist()))
        with fc3:
            lvl_filter  = st.selectbox("Occupation Level", ["All","Trainee","Starter","Competent","Master"])

        display_mh = mh.copy()
        if dept_filter != "All":
            display_mh = display_mh[display_mh["Department"] == dept_filter]
        if comp_filter != "All":
            display_mh = display_mh[display_mh["Company Name"] == comp_filter]
        if lvl_filter != "All":
            display_mh = display_mh[display_mh["Occupation Level"] == lvl_filter]

        # Merge UPH
        if not uph.empty:
            uph_pivot = uph.groupby("username").agg(UPH=("UPH","mean"), Report_Type=("Report Type", lambda x: ", ".join(x.unique()))).reset_index()
            uph_pivot["UPH"] = uph_pivot["UPH"].round(1)
            display_mh = display_mh.merge(uph_pivot, left_on="Cust-Oracle Username", right_on="username", how="left")

        # Table columns
        show_cols = ["Cust-Oracle Username","Name Surname","Company Name","Department","Occupation","Occupation Level","Weeks On Site","Hire Date"]
        if "UPH" in display_mh.columns:
            show_cols += ["UPH","Report_Type"]

        # Search
        search = st.text_input("🔍 Search by name or username")
        if search:
            display_mh = display_mh[
                display_mh["Name Surname"].astype(str).str.lower().str.contains(search.lower()) |
                display_mh["Cust-Oracle Username"].astype(str).str.lower().str.contains(search.lower())
            ]

        st.markdown(f"**Showing {len(display_mh)} employees**")
        final_table = display_mh[show_cols].rename(columns={
            "Cust-Oracle Username": "Username",
            "Name Surname": "Full Name",
            "Company Name": "Company",
            "Occupation Level": "Level",
            "Weeks On Site": "Weeks",
            "Hire Date": "Hire Date",
            "Report_Type": "Role(s)"
        }).reset_index(drop=True)
        final_table["Hire Date"] = pd.to_datetime(final_table["Hire Date"], errors="coerce").dt.strftime("%Y-%m-%d")

        st.dataframe(final_table, use_container_width=True, hide_index=True)

        # ── User drill-down ──
        st.markdown("---")
        st.markdown('<div class="section-header">🔎 Individual User Detail</div>', unsafe_allow_html=True)

        username_input = st.text_input("Enter Username (Oracle) to see full detail:")
        if username_input:
            user_mh = mh[mh["Cust-Oracle Username"].astype(str).str.lower().str.strip() == username_input.lower().strip()]
            if user_mh.empty:
                st.warning(f"No employee found with username **{username_input}**")
            else:
                row = user_mh.iloc[0]
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

                # Daily hours
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
                c2.metric("Avg Daily Hours", f"{total_hrs/7:.1f} hrs")

                # UPH for this user
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

        dept_sel = st.selectbox("Filter by Department", ["All"] + sorted(mh["Department"].dropna().unique().tolist()), key="daily_dept")
        disp = mh.copy()
        if dept_sel != "All":
            disp = disp[disp["Department"] == dept_sel]

        # Build hours table
        records = []
        for _, r in disp.iterrows():
            row_data = {"Username": r.get("Cust-Oracle Username",""), "Name": r.get("Name Surname",""),
                        "Company": r.get("Company Name",""), "Level": r.get("Occupation Level","")}
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

        # Summary
        st.markdown('<div class="section-header">📊 Daily Total Hours Across Team</div>', unsafe_allow_html=True)
        daily_totals = []
        for d in day_labels:
            if d in hours_table.columns:
                daily_totals.append({"Day": d, "Total Hours": hours_table[d].sum()})
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
        st.info("Upload rate reports (picking/packing/putaway/receiving) and process them to see the leaderboard.")
    else:
        rtype_sel = st.selectbox("Report Type", ["All"] + sorted(uph["Report Type"].unique().tolist()))
        disp_uph = uph.copy()
        if rtype_sel != "All":
            disp_uph = disp_uph[disp_uph["Report Type"] == rtype_sel]

        disp_uph = disp_uph.sort_values("UPH", ascending=False).reset_index(drop=True)
        disp_uph.index += 1

        show = disp_uph[["Name Surname","username","Company","Department","Occupation","Occupation Level","Weeks On Site","Report Type","Total Units","UPH"]].copy()
        show.columns = ["Name","Username","Company","Department","Occupation","Level","Weeks","Type","Units","UPH"]

        st.dataframe(show, use_container_width=True)

        # Chart: top 20
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
