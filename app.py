
import streamlit as st
import pandas as pd
import zipfile
import io
import altair as alt

st.set_page_config(page_title="DC Performance Dashboard", page_icon="📦", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 20px 30px; border-radius: 12px; margin-bottom: 24px;
    }
    .main-header h1 { margin: 0; font-size: 2rem; color: white; }
    .main-header p  { margin: 4px 0 0; opacity: .85; color: white; }
    .metric-card {
        background: white; border-radius: 10px; padding: 18px 20px;
        border-left: 5px solid #2d6a9f; box-shadow: 0 2px 8px rgba(0,0,0,.08); margin-bottom: 12px;
    }
    .metric-card h3 { margin: 0 0 4px; font-size: .82rem; color: #666; text-transform: uppercase; letter-spacing: .05em; }
    .metric-card p  { margin: 0; font-size: 1.8rem; font-weight: 700; color: #1e3a5f; }
    .metric-green   { border-left-color: #27ae60; }
    .metric-red     { border-left-color: #e74c3c; }
    .metric-orange  { border-left-color: #f39c12; }
    .badge-trainee   { background:#ffeaa7; color:#d35400; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-starter   { background:#a8e6cf; color:#1e8449; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-competent { background:#74b9ff; color:#1a5276; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-master    { background:#fd79a8; color:#6c1837; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .user-detail-card {
        background: #f8fafc; border-radius: 10px; padding: 20px;
        border: 1px solid #e2e8f0; margin-bottom: 16px;
    }
    .section-header {
        background: #eef2f7; padding: 10px 16px; border-radius: 8px;
        font-weight: 700; color: #1e3a5f; margin: 16px 0 10px;
    }
    .status-present { color: #27ae60; font-weight: 700; }
    .status-off     { color: #e74c3c; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# ── Occupation keywords: which job titles belong to each report type ───────────
OCCUPATION_KEYWORDS = {
    "Picking":   ["picker"],
    "Packing":   ["packer", "sealer", "packing"],
    "Receiving": ["receiver", "receiving"],
    "Putaway":   ["putaway", "filler", "sorter", "grid"],
}

def matches_occupation(occupation, report_type):
    """Return True if the employee occupation matches the report type."""
    if not occupation or str(occupation).strip() in ["-", "", "nan"]:
        return False
    occ_lower = str(occupation).lower()
    keywords  = OCCUPATION_KEYWORDS.get(report_type, [])
    return any(kw in occ_lower for kw in keywords)

# ── Helpers ────────────────────────────────────────────────────────────────────
def get_occupation_level(hire_date_raw):
    try:
        if pd.isna(hire_date_raw):
            return "Trainee"
        hire   = pd.to_datetime(hire_date_raw)
        days   = (pd.Timestamp.today() - hire).days
        months = days / 30.44
        if months >= 6:   return "Master"
        elif months >= 3: return "Competent"
        elif months >= 1: return "Starter"
        else:             return "Trainee"
    except:
        return "Trainee"

def badge_html(level):
    lvl = str(level).lower()
    if lvl not in ["trainee","starter","competent","master"]:
        lvl = "trainee"
    return f'<span class="badge-{lvl}">{level}</span>'

def parse_man_hours(f):
    df_raw    = pd.read_excel(f, header=None)
    day_row   = df_raw.iloc[8].tolist()
    col_names = df_raw.iloc[9].tolist()

    cols        = []
    current_day = ""
    day_slots   = ["In", "Out", "Man Hrs", "Missing"]
    day_labels_found = []

    for i, val in enumerate(col_names):
        dv = day_row[i]
        if isinstance(dv, str) and dv.strip():
            current_day = dv.strip()
            if current_day not in day_labels_found and current_day != "Period Totals":
                day_labels_found.append(current_day)
        if isinstance(val, str) and val.strip():
            name = val.strip()
            if name in day_slots and current_day:
                cols.append(f"{current_day}_{name}")
            else:
                cols.append(name)
        else:
            cols.append(f"col_{i}")

    data = df_raw.iloc[10:].copy()
    data.columns = cols[:len(data.columns)]
    data = data[data["Site"].notna() & (data["Site"] != "Site")].copy()
    data = data[data["Name Surname"].notna()].copy()
    data = data[~data["Name Surname"].str.strip().eq("Total")].copy()
    data["Hire Date"]        = pd.to_datetime(data["Hire Date"], errors="coerce")
    data["Occupation Level"] = data["Hire Date"].apply(get_occupation_level)

    def weeks_on_site(hd):
        try:
            if pd.isna(hd): return 0
            return max(0, round((pd.Timestamp.today() - pd.to_datetime(hd)).days / 7, 1))
        except:
            return 0
    data["Weeks On Site"] = data["Hire Date"].apply(weeks_on_site)

    # Parse man hours per day as float
    for d in day_labels_found:
        mh_col = f"{d}_Man Hrs"
        in_col = f"{d}_In"
        if mh_col in data.columns:
            data[f"{d}_hrs"] = pd.to_numeric(data[mh_col], errors="coerce").fillna(0)
        else:
            data[f"{d}_hrs"] = 0.0

    return data, day_labels_found

def parse_rate_csv(f, report_type):
    df_raw     = pd.read_csv(f, header=None)
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
                try:    date_val = pd.to_datetime(cell)
                except: pass
            if "Target" in cell_str and "=" in cell_str:
                target_val = cell_str.strip()

    if header_row is None:
        return pd.DataFrame(), date_val, target_val, []

    hour_row  = df_raw.iloc[header_row - 1].tolist() if header_row > 0 else []
    base_cols = ["username", "Hour", "Total"]
    hour_cols = []
    hour_names = []  # clean hour labels e.g. "2", "3", "14"

    for i in range(3, len(df_raw.columns)):
        hv     = hour_row[i] if i < len(hour_row) else ""
        hv_str = str(hv).replace(".0","").strip()
        if hv_str not in ["", "nan"]:
            hour_cols.append(f"h_{hv_str}")
            hour_names.append(hv_str)
        else:
            hour_cols.append(f"h_col{i}")
            hour_names.append("")

    all_cols = base_cols + hour_cols
    data     = df_raw.iloc[header_row + 1:].copy()
    data.columns = all_cols[:len(data.columns)]
    data = data[data["username"].notna()].copy()
    data = data[~data["username"].astype(str).str.lower().str.strip().eq("total")].copy()
    data["Total"] = pd.to_numeric(data["Total"], errors="coerce").fillna(0)
    for hc in hour_cols:
        if hc in data.columns:
            data[hc] = pd.to_numeric(data[hc], errors="coerce").fillna(0)
    data["report_type"] = report_type
    data["report_date"] = date_val if date_val else pd.Timestamp.today().normalize()
    return data, date_val, target_val, hour_names

def load_zip_csv(uploaded_file, report_type):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
        with z.open(z.namelist()[0]) as f:
            return parse_rate_csv(f, report_type)

def compute_uph(rate_data, man_hours_data):
    """
    Calculate UPH only for employees whose occupation matches the report type.
    Pickers -> Picking, Packers -> Packing, Receivers -> Receiving, Putaway/Fillers -> Putaway
    """
    if rate_data.empty or man_hours_data.empty:
        return pd.DataFrame()

    uph_rows = []
    skipped  = 0

    for _, row in rate_data.iterrows():
        username    = str(row["username"]).strip().lower()
        total_units = row["Total"]
        rt          = row["report_type"]
        rd          = row["report_date"]

        mh_match = man_hours_data[
            man_hours_data["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username
        ]
        if mh_match.empty:
            continue

        mh         = mh_match.iloc[0]
        occupation = mh.get("Occupation", "")

        # Only include if occupation matches report type
        if not matches_occupation(occupation, rt):
            skipped += 1
            continue

        # Get hours worked that day
        day_hrs = next(
            (mh.get(c, 0) for c in man_hours_data.columns if "_hrs" in c and mh.get(c, 0) > 0),
            8.0
        )

        uph_rows.append({
            "username":         username,
            "Name Surname":     mh.get("Name Surname", ""),
            "Company":          mh.get("Company Name", ""),
            "Occupation":       occupation,
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

# ── Session state ──────────────────────────────────────────────────────────────
for key, default in [
    ("man_hours_df",  None),
    ("day_labels",    []),
    ("rate_dfs",      []),
    ("rate_meta",     {}),
    ("uph_df",        pd.DataFrame()),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>DC Performance Dashboard</h1>
  <p>Man Hours | Units Per Hour | Occupation Levels | Daily and Weekly Tracking</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Upload Daily Files")
    mh_file   = st.file_uploader("Auto Man Hours Report (.xlsx)", type=["xlsx"], key="mh")
    pick_file = st.file_uploader("Picking Rate Report (.zip)",    type=["zip"],  key="pick")
    pack_file = st.file_uploader("Packing Rate Report (.zip)",    type=["zip"],  key="pack")
    put_file  = st.file_uploader("Putaway Rate Report (.zip)",    type=["zip"],  key="put")
    recv_file = st.file_uploader("Receiving Rate Report (.zip)",  type=["zip"],  key="recv")
    process_btn = st.button("Process Files", use_container_width=True, type="primary")

    if process_btn:
        with st.spinner("Processing..."):
            errors      = []
            rate_frames = []
            rate_meta   = {}

            if mh_file:
                try:
                    mh_df, day_lbls = parse_man_hours(mh_file)
                    st.session_state.man_hours_df = mh_df
                    st.session_state.day_labels   = day_lbls
                    st.success(f"Man Hours: {len(mh_df)} employees | Days: {len(day_lbls)}")
                except Exception as e:
                    errors.append(f"Man Hours error: {e}")

            for file_obj, rtype in [
                (pick_file,"Picking"),(pack_file,"Packing"),
                (put_file,"Putaway"),(recv_file,"Receiving")
            ]:
                if file_obj:
                    try:
                        df, dv, target, hour_names = load_zip_csv(file_obj, rtype)
                        if not df.empty:
                            rate_frames.append(df)
                            rate_meta[rtype] = {"date": dv, "target": target, "hour_names": hour_names}
                            st.success(f"{rtype}: {len(df)} users")
                    except Exception as e:
                        errors.append(f"{rtype} error: {e}")

            if rate_frames:
                st.session_state.rate_dfs  = rate_frames
                st.session_state.rate_meta = rate_meta
                all_rates = pd.concat(rate_frames, ignore_index=True)
                if st.session_state.man_hours_df is not None:
                    st.session_state.uph_df = compute_uph(all_rates, st.session_state.man_hours_df)
                    st.success(f"UPH calculated for {len(st.session_state.uph_df)} matched workers")

            for e in errors:
                st.error(e)

    st.markdown("---")
    st.markdown("**Occupation matching:**")
    st.markdown("Picking report = Pickers only")
    st.markdown("Packing report = Packers only")
    st.markdown("Receiving report = Receivers only")
    st.markdown("Putaway report = Putaway / Fillers only")
    st.markdown("---")
    st.markdown("**Occupation Levels:**")
    st.markdown("Trainee = up to 2 weeks")
    st.markdown("Starter = more than 1 month")
    st.markdown("Competent = 3 months or more")
    st.markdown("Master = 6 months or more")

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Overview",
    "User Detail",
    "Daily Hours",
    "Leaderboard",
    "Days Rate Detail"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 - OVERVIEW: Show who signed in/out, who was off
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    mh         = st.session_state.man_hours_df
    day_labels = st.session_state.day_labels

    if mh is None:
        st.info("Upload files and click Process Files to get started.")
    else:
        st.markdown('<div class="section-header">Select a Day to View Attendance</div>', unsafe_allow_html=True)

        if not day_labels:
            st.warning("No day columns found in the man hours file.")
        else:
            selected_day = st.selectbox("Day", day_labels, key="overview_day")
            in_col  = f"{selected_day}_In"
            out_col = f"{selected_day}_Out"
            mh_col  = f"{selected_day}_Man Hrs"

            # Split present vs off
            if in_col in mh.columns:
                present_mask = mh[in_col].notna() & (mh[in_col].astype(str).str.strip() != "") & (mh[in_col].astype(str).str.strip() != "nan")
                present_df   = mh[present_mask].copy()
                absent_df    = mh[~present_mask].copy()
            else:
                present_df = pd.DataFrame()
                absent_df  = mh.copy()

            # KPI row
            uph = st.session_state.uph_df
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f'<div class="metric-card metric-green"><h3>Signed In Today</h3><p>{len(present_df)}</p></div>', unsafe_allow_html=True)
            with k2:
                st.markdown(f'<div class="metric-card metric-red"><h3>Absent / Off</h3><p>{len(absent_df)}</p></div>', unsafe_allow_html=True)
            with k3:
                total_hrs = 0
                if not present_df.empty and mh_col in present_df.columns:
                    hrs_col = f"{selected_day}_hrs"
                    if hrs_col in present_df.columns:
                        total_hrs = present_df[hrs_col].sum()
                st.markdown(f'<div class="metric-card"><h3>Total Hours Clocked</h3><p>{total_hrs:.0f} hrs</p></div>', unsafe_allow_html=True)
            with k4:
                avg_uph = uph["UPH"].mean() if not uph.empty else 0
                st.markdown(f'<div class="metric-card metric-orange"><h3>Avg UPH (Matched)</h3><p>{avg_uph:.1f}</p></div>', unsafe_allow_html=True)

            # Present employees table
            st.markdown(f'<div class="section-header">Signed In - {selected_day} ({len(present_df)} employees)</div>', unsafe_allow_html=True)
            if not present_df.empty:
                show_present = present_df[[
                    "Cust-Oracle Username", "Name Surname", "Company Name",
                    "Department", "Occupation", "Occupation Level", "Weeks On Site"
                ]].copy()
                if in_col in present_df.columns:
                    show_present["Sign In"]  = present_df[in_col].astype(str)
                if out_col in out_col and out_col in present_df.columns:
                    show_present["Sign Out"] = present_df[out_col].astype(str)
                hrs_col = f"{selected_day}_hrs"
                if hrs_col in present_df.columns:
                    show_present["Hours Worked"] = present_df[hrs_col].round(2)
                show_present = show_present.rename(columns={
                    "Cust-Oracle Username": "Username",
                    "Name Surname":        "Full Name",
                    "Company Name":        "Company",
                    "Occupation Level":    "Level",
                    "Weeks On Site":       "Weeks"
                })
                # Filter
                dept_filter_ov = st.selectbox("Filter by Department",
                    ["All"] + sorted(mh["Department"].dropna().unique().tolist()), key="ov_dept")
                if dept_filter_ov != "All":
                    show_present = show_present[show_present["Department"] == dept_filter_ov]

                st.dataframe(show_present.reset_index(drop=True), use_container_width=True, hide_index=True)

                # Chart: hours worked by department
                if hrs_col in present_df.columns:
                    dept_hrs = present_df.groupby("Department")[hrs_col].sum().reset_index()
                    dept_hrs.columns = ["Department","Hours"]
                    dept_hrs = dept_hrs[dept_hrs["Hours"] > 0].sort_values("Hours", ascending=False)
                    if not dept_hrs.empty:
                        st.markdown('<div class="section-header">Hours by Department</div>', unsafe_allow_html=True)
                        ch = alt.Chart(dept_hrs).mark_bar(
                            cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#2d6a9f"
                        ).encode(
                            x=alt.X("Hours:Q"),
                            y=alt.Y("Department:N", sort="-x"),
                            tooltip=["Department","Hours"]
                        ).properties(height=300)
                        st.altair_chart(ch, use_container_width=True)

            # Absent employees
            st.markdown(f'<div class="section-header">Off / Not Signed In - {selected_day} ({len(absent_df)} employees)</div>', unsafe_allow_html=True)
            if not absent_df.empty:
                show_absent = absent_df[[
                    "Cust-Oracle Username","Name Surname","Company Name",
                    "Department","Occupation","Occupation Level","Weeks On Site"
                ]].rename(columns={
                    "Cust-Oracle Username":"Username","Name Surname":"Full Name",
                    "Company Name":"Company","Occupation Level":"Level","Weeks On Site":"Weeks"
                }).reset_index(drop=True)
                st.dataframe(show_absent, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 - USER DETAIL
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">User Detail</div>', unsafe_allow_html=True)
    mh  = st.session_state.man_hours_df
    uph = st.session_state.uph_df
    day_labels = st.session_state.day_labels

    if mh is None:
        st.info("Upload files and process them first.")
    else:
        fc1, fc2, fc3 = st.columns(3)
        with fc1: dept_filter = st.selectbox("Department", ["All"] + sorted(mh["Department"].dropna().unique().tolist()))
        with fc2: comp_filter = st.selectbox("Company",    ["All"] + sorted(mh["Company Name"].dropna().unique().tolist()))
        with fc3: lvl_filter  = st.selectbox("Level", ["All","Trainee","Starter","Competent","Master"])

        display_mh = mh.copy()
        if dept_filter != "All": display_mh = display_mh[display_mh["Department"] == dept_filter]
        if comp_filter != "All": display_mh = display_mh[display_mh["Company Name"] == comp_filter]
        if lvl_filter  != "All": display_mh = display_mh[display_mh["Occupation Level"] == lvl_filter]

        if not uph.empty:
            uph_pivot = uph.groupby("username").agg(
                UPH=("UPH","mean"),
                Report_Type=("Report Type", lambda x: ", ".join(x.unique()))
            ).reset_index()
            uph_pivot["UPH"] = uph_pivot["UPH"].round(1)
            display_mh = display_mh.merge(uph_pivot, left_on="Cust-Oracle Username", right_on="username", how="left")

        show_cols = ["Cust-Oracle Username","Name Surname","Company Name","Department",
                     "Occupation","Occupation Level","Weeks On Site","Hire Date"]
        if "UPH" in display_mh.columns:
            show_cols += ["UPH","Report_Type"]

        search = st.text_input("Search by name or username")
        if search:
            display_mh = display_mh[
                display_mh["Name Surname"].astype(str).str.lower().str.contains(search.lower()) |
                display_mh["Cust-Oracle Username"].astype(str).str.lower().str.contains(search.lower())
            ]

        st.markdown(f"**{len(display_mh)} employees**")
        final_table = display_mh[show_cols].rename(columns={
            "Cust-Oracle Username":"Username","Name Surname":"Full Name",
            "Company Name":"Company","Occupation Level":"Level",
            "Weeks On Site":"Weeks","Report_Type":"Role(s)"
        }).reset_index(drop=True)
        final_table["Hire Date"] = pd.to_datetime(final_table["Hire Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        st.dataframe(final_table, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown('<div class="section-header">Individual User Detail</div>', unsafe_allow_html=True)
        username_input = st.text_input("Enter Username to see full profile:")
        if username_input:
            user_mh = mh[mh["Cust-Oracle Username"].astype(str).str.lower().str.strip() == username_input.lower().strip()]
            if user_mh.empty:
                st.warning(f"No employee found with username: {username_input}")
            else:
                row   = user_mh.iloc[0]
                level = row.get("Occupation Level", "Trainee")
                st.markdown(f"""
                <div class="user-detail-card">
                  <h2 style="margin:0 0 8px;">{row.get("Name Surname","")}</h2>
                  <p><b>Username:</b> {row.get("Cust-Oracle Username","")}</p>
                  <p><b>Company:</b> {row.get("Company Name","")}</p>
                  <p><b>Department:</b> {row.get("Department","")}</p>
                  <p><b>Occupation:</b> {row.get("Occupation","")}</p>
                  <p><b>Hire Date:</b> {str(row.get("Hire Date",""))[:10]}</p>
                  <p><b>Years of Service:</b> {row.get("Years of Service (Yrs/Mths)","")}</p>
                  <p><b>Weeks on Site:</b> {row.get("Weeks On Site", 0)} &nbsp; <b>Level:</b> {badge_html(level)}</p>
                </div>
                """, unsafe_allow_html=True)

                if day_labels:
                    hrs_data  = []
                    total_hrs = 0
                    for d in day_labels:
                        col = f"{d}_hrs"
                        v   = row.get(col, 0) if col in row.index else 0
                        if pd.isna(v): v = 0
                        hrs_data.append({"Day": d, "Hours": float(v)})
                        total_hrs += float(v)
                    hrs_df = pd.DataFrame(hrs_data)

                    st.markdown('<div class="section-header">Hours Worked This Week</div>', unsafe_allow_html=True)
                    hr_chart = alt.Chart(hrs_df).mark_bar(
                        cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#2d6a9f"
                    ).encode(x="Day:N", y="Hours:Q", tooltip=["Day","Hours"]).properties(height=220)
                    st.altair_chart(hr_chart, use_container_width=True)
                    c1, c2 = st.columns(2)
                    n_days = len(day_labels) if day_labels else 7
                    c1.metric("Total Hours This Week", f"{total_hrs:.1f} hrs")
                    c2.metric("Avg Daily Hours",       f"{total_hrs/n_days:.1f} hrs")

                if not uph.empty:
                    user_uph = uph[uph["username"].str.lower() == username_input.lower()]
                    if not user_uph.empty:
                        st.markdown('<div class="section-header">Units Per Hour Performance</div>', unsafe_allow_html=True)
                        uph_show = user_uph[["Report Type","Total Units","Hours Worked","UPH","Report Date"]].copy()
                        uph_show["Report Date"] = pd.to_datetime(uph_show["Report Date"]).dt.strftime("%Y-%m-%d")
                        st.dataframe(uph_show, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 - DAILY HOURS
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Daily and Weekly Hours Tracker</div>', unsafe_allow_html=True)
    mh         = st.session_state.man_hours_df
    day_labels = st.session_state.day_labels

    if mh is None:
        st.info("Upload files and process them first.")
    else:
        dept_sel = st.selectbox("Filter by Department",
            ["All"] + sorted(mh["Department"].dropna().unique().tolist()), key="daily_dept")
        disp = mh.copy()
        if dept_sel != "All":
            disp = disp[disp["Department"] == dept_sel]

        records = []
        for _, r in disp.iterrows():
            row_data = {
                "Username": r.get("Cust-Oracle Username",""),
                "Name":     r.get("Name Surname",""),
                "Company":  r.get("Company Name",""),
                "Level":    r.get("Occupation Level","")
            }
            wk_total = 0
            for d in day_labels:
                col = f"{d}_hrs"
                v   = r.get(col, 0) if col in r.index else 0
                if pd.isna(v): v = 0
                row_data[d] = round(float(v), 2)
                wk_total   += float(v)
            row_data["Weekly Total"] = round(wk_total, 2)
            records.append(row_data)

        hours_table = pd.DataFrame(records)
        hours_table = hours_table[hours_table["Weekly Total"] > 0].sort_values("Weekly Total", ascending=False)
        st.markdown(f"**{len(hours_table)} employees with recorded hours**")
        st.dataframe(hours_table, use_container_width=True, hide_index=True)

        if day_labels:
            st.markdown('<div class="section-header">Daily Total Hours Across Team</div>', unsafe_allow_html=True)
            daily_totals = [
                {"Day": d, "Total Hours": hours_table[d].sum()}
                for d in day_labels if d in hours_table.columns
            ]
            dt_df = pd.DataFrame(daily_totals)
            day_chart = alt.Chart(dt_df).mark_bar(
                cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#27ae60"
            ).encode(x="Day:N", y="Total Hours:Q", tooltip=["Day","Total Hours"]).properties(height=260)
            st.altair_chart(day_chart, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 - LEADERBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">UPH Leaderboard (Occupation Matched)</div>', unsafe_allow_html=True)
    uph = st.session_state.uph_df

    if uph.empty:
        st.info("Upload rate reports and process them to see the leaderboard.")
        st.markdown("""
        **Note:** UPH is calculated only for matched occupations:
        - Picking report = Pickers only
        - Packing report = Packers only
        - Receiving report = Receivers only
        - Putaway report = Putaway / Fillers only
        """)
    else:
        rtype_sel = st.selectbox("Report Type", ["All"] + sorted(uph["Report Type"].unique().tolist()))
        disp_uph  = uph.copy()
        if rtype_sel != "All":
            disp_uph = disp_uph[disp_uph["Report Type"] == rtype_sel]
        disp_uph = disp_uph.sort_values("UPH", ascending=False).reset_index(drop=True)
        disp_uph.index += 1
        show = disp_uph[[
            "Name Surname","username","Company","Department","Occupation",
            "Occupation Level","Weeks On Site","Report Type","Total Units","UPH"
        ]].copy()
        show.columns = ["Name","Username","Company","Department","Occupation","Level","Weeks","Type","Units","UPH"]
        st.dataframe(show, use_container_width=True)

        top20    = disp_uph.head(20)
        lb_chart = alt.Chart(top20).mark_bar(
            cornerRadiusTopLeft=5, cornerRadiusTopRight=5
        ).encode(
            x=alt.X("UPH:Q"),
            y=alt.Y("Name Surname:N", sort="-x"),
            color=alt.Color("Occupation Level:N",
                scale=alt.Scale(
                    domain=["Trainee","Starter","Competent","Master"],
                    range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8"]
                )),
            tooltip=["Name Surname","UPH","Report Type","Occupation","Company"]
        ).properties(height=500, title="Top 20 Workers - Units Per Hour")
        st.altair_chart(lb_chart, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 - DAYS RATE DETAIL (same layout as CSV + employee details)
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-header">Days Rate Detail - Hourly Breakdown Per User</div>', unsafe_allow_html=True)

    mh        = st.session_state.man_hours_df
    rate_dfs  = st.session_state.rate_dfs
    rate_meta = st.session_state.rate_meta

    if not rate_dfs:
        st.info("Upload at least one rate report and click Process Files.")
    else:
        available_types = list(rate_meta.keys())
        selected_type   = st.selectbox("Select Report Type", available_types)
        rate_df         = next((d for d in rate_dfs if d["report_type"].iloc[0] == selected_type), None)

        if rate_df is None:
            st.warning("No data for this report type.")
        else:
            meta        = rate_meta.get(selected_type, {})
            report_date = meta.get("date")
            target_str  = meta.get("target", "")
            hour_names  = meta.get("hour_names", [])
            date_str    = report_date.strftime("%d %B %Y") if report_date else "N/A"

            # Info banner
            i1, i2, i3 = st.columns(3)
            with i1: st.markdown(f'<div class="metric-card"><h3>Report Type</h3><p style="font-size:1.2rem">{selected_type}</p></div>', unsafe_allow_html=True)
            with i2: st.markdown(f'<div class="metric-card"><h3>Report Date</h3><p style="font-size:1.2rem">{date_str}</p></div>', unsafe_allow_html=True)
            with i3: st.markdown(f'<div class="metric-card"><h3>Target</h3><p style="font-size:1rem">{target_str if target_str else "N/A"}</p></div>', unsafe_allow_html=True)

            # Hour columns from the dataframe
            hour_cols = [c for c in rate_df.columns if c.startswith("h_") and not c.startswith("h_col")]

            # Filters
            f1, f2, f3 = st.columns(3)
            with f1: search_user = st.text_input("Search Username or Name", key="rate_search")
            with f2:
                dept_opts = ["All"]
                comp_opts = ["All"]
                if mh is not None:
                    dept_opts += sorted(mh["Department"].dropna().unique().tolist())
                    comp_opts += sorted(mh["Company Name"].dropna().unique().tolist())
                dept_f = st.selectbox("Department", dept_opts, key="rate_dept")
            with f3:
                comp_f = st.selectbox("Company", comp_opts, key="rate_comp")

            # ── Build table - same layout as CSV + details ───────────────────
            records = []
            for _, row in rate_df.iterrows():
                username    = str(row["username"]).strip()
                total_units = row["Total"]

                # Defaults
                full_name = username; company = "-"; dept = "-"
                occupation = "-"; hire_date = "-"
                weeks = 0; level = "-"; years_svc = "-"

                if mh is not None:
                    mh_match = mh[mh["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username.lower()]
                    if not mh_match.empty:
                        mhr        = mh_match.iloc[0]
                        full_name  = mhr.get("Name Surname", username)
                        company    = mhr.get("Company Name", "-")
                        dept       = mhr.get("Department", "-")
                        occupation = mhr.get("Occupation", "-")
                        hire_date  = str(mhr.get("Hire Date", ""))[:10]
                        weeks      = mhr.get("Weeks On Site", 0)
                        level      = mhr.get("Occupation Level", "Trainee")
                        years_svc  = mhr.get("Years of Service (Yrs/Mths)", "-")

                rec = {
                    "Username":         username,
                    "Full Name":        full_name,
                    "Company":          company,
                    "Department":       dept,
                    "Occupation":       occupation,
                    "Level":            level,
                    "Weeks On Site":    weeks,
                    "Years of Service": years_svc,
                    "Hire Date":        hire_date,
                    "Total":            int(total_units),
                }

                # Add hour columns exactly as they appear in CSV (Hr 2, Hr 3 ... Hr 14)
                for hc, hn in zip(hour_cols, hour_names):
                    if hn and hn != "":
                        col_label = f"Hr {hn}"
                    else:
                        col_label = hc
                    val = row.get(hc, 0)
                    rec[col_label] = int(val) if not pd.isna(val) else 0

                # Calculated fields
                active_hours  = sum(1 for hc in hour_cols if row.get(hc, 0) > 0)
                rec["Active Hrs"] = active_hours
                rec["Avg UPH"]    = round(total_units / active_hours, 1) if active_hours > 0 else 0

                records.append(rec)

            result_df = pd.DataFrame(records)

            # Apply filters
            if search_user:
                result_df = result_df[
                    result_df["Username"].str.lower().str.contains(search_user.lower()) |
                    result_df["Full Name"].str.lower().str.contains(search_user.lower())
                ]
            if dept_f != "All":
                result_df = result_df[result_df["Department"] == dept_f]
            if comp_f != "All":
                result_df = result_df[result_df["Company"] == comp_f]

            result_df = result_df.sort_values("Total", ascending=False).reset_index(drop=True)

            # Summary KPIs
            st.markdown('<div class="section-header">Summary</div>', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Workers Active",  len(result_df))
            k2.metric("Total Units",     f"{result_df['Total'].sum():,}")
            k3.metric("Avg UPH",         f"{result_df['Avg UPH'].mean():.1f}")
            k4.metric("Best UPH",        f"{result_df['Avg UPH'].max():.1f}")

            # Full table (same layout as CSV + details)
            st.markdown(f'<div class="section-header">All Workers - {selected_type} | {date_str}</div>', unsafe_allow_html=True)
            st.markdown(f"**{len(result_df)} workers**")
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            # ── Individual drill-down ─────────────────────────────────────────
            st.markdown("---")
            st.markdown('<div class="section-header">Select a Worker - Hourly Breakdown</div>', unsafe_allow_html=True)
            selected_u = st.selectbox("Select worker:", ["-- Select --"] + result_df["Username"].tolist(), key="rate_user_select")

            if selected_u != "-- Select --":
                user_row  = result_df[result_df["Username"] == selected_u].iloc[0]
                level_val = str(user_row.get("Level","Trainee"))

                st.markdown(f"""
                <div class="user-detail-card">
                  <h2 style="margin:0 0 8px;">{user_row["Full Name"]}</h2>
                  <div style="display:flex; flex-wrap:wrap; gap:20px; margin-top:8px;">
                    <div><b>Username:</b> {user_row["Username"]}</div>
                    <div><b>Company:</b> {user_row["Company"]}</div>
                    <div><b>Department:</b> {user_row["Department"]}</div>
                    <div><b>Occupation:</b> {user_row["Occupation"]}</div>
                    <div><b>Hire Date:</b> {user_row["Hire Date"]}</div>
                    <div><b>Years of Service:</b> {user_row["Years of Service"]}</div>
                    <div><b>Weeks on Site:</b> {user_row["Weeks On Site"]}</div>
                    <div><b>Level:</b> {badge_html(level_val)}</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                hour_label_cols = [c for c in result_df.columns if c.startswith("Hr ")]
                if hour_label_cols:
                    st.markdown('<div class="section-header">Units Produced Per Hour</div>', unsafe_allow_html=True)
                    hour_data = [{"Hour": hc, "Units": int(user_row.get(hc, 0))} for hc in hour_label_cols]
                    hour_df   = pd.DataFrame(hour_data)

                    uh1, uh2, uh3, uh4 = st.columns(4)
                    uh1.metric("Total Units",  f"{int(user_row['Total']):,}")
                    uh2.metric("Active Hours", user_row["Active Hrs"])
                    uh3.metric("Avg UPH",      user_row["Avg UPH"])
                    non_zero = hour_df[hour_df["Units"] > 0]
                    peak_h   = non_zero.loc[non_zero["Units"].idxmax(), "Hour"] if not non_zero.empty else "-"
                    uh4.metric("Peak Hour", peak_h)

                    q66 = hour_df["Units"].quantile(0.66)
                    q33 = hour_df["Units"].quantile(0.33)
                    bar = alt.Chart(hour_df).mark_bar(
                        cornerRadiusTopLeft=5, cornerRadiusTopRight=5
                    ).encode(
                        x=alt.X("Hour:N", sort=None, title="Hour of Day"),
                        y=alt.Y("Units:Q", title="Units Produced"),
                        color=alt.condition(
                            alt.datum.Units >= q66,
                            alt.value("#27ae60"),
                            alt.condition(
                                alt.datum.Units >= q33,
                                alt.value("#f39c12"),
                                alt.value("#e74c3c")
                            )
                        ),
                        tooltip=["Hour","Units"]
                    ).properties(height=300, title=f"Hourly Units - {user_row['Full Name']} ({selected_type})")
                    st.altair_chart(bar, use_container_width=True)

                    st.dataframe(non_zero.reset_index(drop=True), use_container_width=True, hide_index=True)

            # All workers UPH chart
            st.markdown('<div class="section-header">Units Per Hour - All Workers</div>', unsafe_allow_html=True)
            chart_data = result_df[result_df["Avg UPH"] > 0].sort_values("Avg UPH", ascending=False).head(30)
            if not chart_data.empty:
                uph_bar = alt.Chart(chart_data).mark_bar(
                    cornerRadiusTopLeft=5, cornerRadiusTopRight=5
                ).encode(
                    x=alt.X("Avg UPH:Q", title="Units Per Hour"),
                    y=alt.Y("Full Name:N", sort="-x", title="Worker"),
                    color=alt.Color("Level:N",
                        scale=alt.Scale(
                            domain=["Trainee","Starter","Competent","Master","-"],
                            range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8","#dfe6e9"]
                        )),
                    tooltip=["Full Name","Username","Company","Avg UPH","Total","Level","Active Hrs"]
                ).properties(height=600, title=f"Top 30 by UPH - {selected_type}")
                st.altair_chart(uph_bar, use_container_width=True)
