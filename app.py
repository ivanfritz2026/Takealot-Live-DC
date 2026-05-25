import streamlit as st
import pandas as pd
import zipfile
import io
import altair as alt

st.set_page_config(page_title="DC Performance Dashboard", page_icon="📦", layout="wide")

st.markdown("""
<style>
    .main-header { background:linear-gradient(135deg,#1e3a5f 0%,#2d6a9f 100%); padding:20px 30px; border-radius:12px; margin-bottom:24px; }
    .main-header h1 { margin:0; font-size:2rem; color:white; }
    .main-header p  { margin:4px 0 0; opacity:.85; color:white; }
    .metric-card { background:white; border-radius:10px; padding:18px 20px; border-left:5px solid #2d6a9f; box-shadow:0 2px 8px rgba(0,0,0,.08); margin-bottom:12px; }
    .metric-card h3 { margin:0 0 4px; font-size:.78rem; color:#666; text-transform:uppercase; letter-spacing:.05em; }
    .metric-card p  { margin:0; font-size:1.7rem; font-weight:700; color:#1e3a5f; }
    .mc-green  { border-left-color:#27ae60; }
    .mc-red    { border-left-color:#e74c3c; }
    .mc-orange { border-left-color:#f39c12; }
    .mc-yellow { border-left-color:#f1c40f; }
    .mc-teal   { border-left-color:#1abc9c; }
    .mc-blue   { border-left-color:#3498db; }
    .mc-pink   { border-left-color:#e91e8c; }
    .badge-trainee   { background:#ffeaa7; color:#d35400; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-starter   { background:#a8e6cf; color:#1e8449; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-competent { background:#74b9ff; color:#1a5276; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .badge-master    { background:#fd79a8; color:#6c1837; padding:3px 10px; border-radius:12px; font-weight:600; font-size:.82rem; }
    .user-detail-card { background:#f8fafc; border-radius:10px; padding:20px; border:1px solid #e2e8f0; margin-bottom:16px; }
    .section-header   { background:#eef2f7; padding:10px 16px; border-radius:8px; font-weight:700; color:#1e3a5f; margin:16px 0 10px; }
</style>
""", unsafe_allow_html=True)

OCCUPATION_KEYWORDS = {
    "Picking":   ["picker"],
    "Packing":   ["packer", "sealer", "packing"],
    "Receiving": ["receiver", "receiving"],
    "Putaway":   ["putaway", "filler", "sorter", "grid"],
}

LEVEL_COLORS = {
    "Trainee":   "background-color:#ffeaa7; color:#d35400; font-weight:600;",
    "Starter":   "background-color:#a8e6cf; color:#1e8449; font-weight:600;",
    "Competent": "background-color:#74b9ff; color:#1a5276; font-weight:600;",
    "Master":    "background-color:#fd79a8; color:#6c1837; font-weight:600;",
}

def color_level_col(val):
    return LEVEL_COLORS.get(str(val), "")

def style_with_levels(df):
    if "Level" in df.columns:
        return df.style.map(color_level_col, subset=["Level"])
    return df.style

def matches_occupation(occupation, report_type):
    if not occupation or str(occupation).strip() in ["-","","nan"]:
        return False
    return any(kw in str(occupation).lower() for kw in OCCUPATION_KEYWORDS.get(report_type, []))

def get_occupation_level(hire_date_raw):
    try:
        if pd.isna(hire_date_raw): return "Trainee"
        months = (pd.Timestamp.today() - pd.to_datetime(hire_date_raw)).days / 30.44
        if months >= 6:   return "Master"
        elif months >= 3: return "Competent"
        elif months >= 1: return "Starter"
        else:             return "Trainee"
    except:
        return "Trainee"

def badge_html(level):
    lvl = str(level).lower()
    if lvl not in ["trainee","starter","competent","master"]: lvl = "trainee"
    return f'<span class="badge-{lvl}">{level}</span>'

def parse_man_hours(f):
    df_raw    = pd.read_excel(f, header=None)
    day_row   = df_raw.iloc[8].tolist()
    col_names = df_raw.iloc[9].tolist()
    cols = []; current_day = ""; day_labels_found = []
    day_slots = ["In","Out","Man Hrs","Missing"]

    for i, val in enumerate(col_names):
        dv = day_row[i]
        # KEY FIX: always update current_day including "Period Totals"
        # This prevents duplicate column names for the last day
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
    n    = min(len(data.columns), len(cols))
    data = data.iloc[:, :n]
    data.columns = cols[:n]

    data = data[data["Site"].notna() & (data["Site"] != "Site")].copy()
    data = data[data["Name Surname"].notna()].copy()
    data = data[~data["Name Surname"].astype(str).str.strip().eq("Total")].copy()

    data["Hire Date"]        = pd.to_datetime(data["Hire Date"], errors="coerce")
    data["Occupation Level"] = data["Hire Date"].apply(get_occupation_level)

    def weeks_on_site(hd):
        try:
            # hd is already a Timestamp after conversion above
            if pd.isna(hd): return 0
            return max(0, round((pd.Timestamp.today() - hd).days / 7, 1))
        except:
            return 0

    data["Weeks On Site"] = data["Hire Date"].apply(weeks_on_site)

    for d in day_labels_found:
        mh_col = f"{d}_Man Hrs"
        if mh_col in data.columns:
            # Man Hrs column contains time strings like "8:30:00" not numbers
            # Convert time string to float hours
            def time_str_to_hrs(val):
                try:
                    if pd.isna(val): return 0.0
                    s = str(val).strip()
                    if ":" in s:
                        parts = s.split(":")
                        return round(int(parts[0]) + int(parts[1])/60, 2)
                    return float(s)
                except:
                    return 0.0
            data[f"{d}_hrs"] = data[mh_col].apply(time_str_to_hrs)
        else:
            data[f"{d}_hrs"] = 0.0

    return data, day_labels_found

def parse_rate_csv(f, report_type):
    df_raw = pd.read_csv(f, header=None)
    header_row = None; date_val = None; target_val = None

    for i, row in df_raw.iterrows():
        r = row.tolist()
        if str(r[0]).strip().lower() in ["user","create_user"]:
            header_row = i
        for cell in r:
            cs = str(cell).strip()
            if "Target" in cs and "=" in cs:
                target_val = cs
            if len(cs) == 10 and cs.count("/") == 2:
                try: date_val = pd.to_datetime(cs)
                except: pass

    if header_row is None:
        return pd.DataFrame(), date_val, target_val, []

    hdr = df_raw.iloc[header_row].tolist()
    col2_str = str(hdr[2]).replace(".0","").strip()
    try:
        int(float(col2_str))
        col2_is_hour = True
    except:
        col2_is_hour = False

    hour_col_map = {}; total_col_idx = None

    if col2_is_hour:
        for i in range(2, len(hdr)):
            hs = str(hdr[i]).replace(".0","").strip()
            try:
                hour_col_map[i] = str(int(float(hs)))
            except:
                if "total" in hs.lower() and total_col_idx is None:
                    total_col_idx = i
    else:
        total_col_idx = 2
        for i in range(3, len(hdr)):
            hs = str(hdr[i]).replace(".0","").strip()
            try:
                hour_col_map[i] = str(int(float(hs)))
            except:
                pass

    all_cols = ["username","_label"]
    for i in range(2, len(hdr)):
        if i in hour_col_map:
            all_cols.append(f"h_{hour_col_map[i]}")
        elif i == total_col_idx:
            all_cols.append("Total")
        else:
            all_cols.append(f"_skip_{i}")

    data = df_raw.iloc[header_row + 1:].copy()
    n    = min(len(data.columns), len(all_cols))
    data = data.iloc[:, :n]
    data.columns = all_cols[:n]

    data = data[data["username"].notna()].copy()
    data = data[~data["username"].astype(str).str.lower().str.strip().isin(["total","nan"])].copy()

    hour_cols = [c for c in data.columns if c.startswith("h_")]
    for hc in hour_cols:
        data[hc] = pd.to_numeric(data[hc], errors="coerce").fillna(0)

    data["Total"]       = data[hour_cols].sum(axis=1)
    data["_active_hrs"] = (data[hour_cols] > 0).sum(axis=1)
    data["report_type"] = report_type
    data["report_date"] = date_val if date_val else pd.Timestamp.today().normalize()

    sorted_hours = [hour_col_map[k] for k in sorted(hour_col_map.keys())]
    return data, date_val, target_val, sorted_hours

def load_zip_csv(uploaded_file, report_type):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
        with z.open(z.namelist()[0]) as f:
            return parse_rate_csv(f, report_type)

def compute_uph(rate_data, man_hours_data):
    if rate_data.empty or man_hours_data.empty:
        return pd.DataFrame()
    rows = []
    for _, row in rate_data.iterrows():
        username    = str(row["username"]).strip().lower()
        total_units = row["Total"]
        rt          = row["report_type"]
        rd          = row["report_date"]
        active_hrs  = max(int(row.get("_active_hrs", 1)), 1)
        mm = man_hours_data[man_hours_data["Cust-Oracle Username"].astype(str).str.strip().str.lower() == username]
        if mm.empty: continue
        mhr = mm.iloc[0]; occupation = mhr.get("Occupation","")
        if not matches_occupation(occupation, rt): continue
        rows.append({
            "username":         username,
            "Name Surname":     mhr.get("Name Surname",""),
            "Company":          mhr.get("Company Name",""),
            "Occupation":       occupation,
            "Department":       mhr.get("Department",""),
            "Hire Date":        mhr.get("Hire Date",""),
            "Weeks On Site":    mhr.get("Weeks On Site",0),
            "Occupation Level": mhr.get("Occupation Level","Trainee"),
            "Report Type":      rt, "Report Date": rd,
            "Total Units":      total_units, "Active Hours": active_hrs,
            "UPH":              round(total_units / active_hrs, 1),
        })
    return pd.DataFrame(rows)

for key, default in [("man_hours_df",None),("day_labels",[]),("rate_dfs",[]),("rate_meta",{}),("uph_df",pd.DataFrame())]:
    if key not in st.session_state: st.session_state[key] = default

st.markdown("""
<div class="main-header">
  <h1>DC Performance Dashboard</h1>
  <p>Man Hours | Units Per Hour | Occupation Levels | Daily Tracking</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Upload Daily Files")
    mh_file   = st.file_uploader("Auto Man Hours (.xlsx)", type=["xlsx"], key="mh")
    pick_file = st.file_uploader("Picking Rate (.zip)",    type=["zip"],  key="pick")
    pack_file = st.file_uploader("Packing Rate (.zip)",    type=["zip"],  key="pack")
    put_file  = st.file_uploader("Putaway Rate (.zip)",    type=["zip"],  key="put")
    recv_file = st.file_uploader("Receiving Rate (.zip)",  type=["zip"],  key="recv")
    process_btn = st.button("Process Files", use_container_width=True, type="primary")

    if process_btn:
        with st.spinner("Processing..."):
            errors=[]; rate_frames=[]; rate_meta={}
            if mh_file:
                try:
                    mh_df, day_lbls = parse_man_hours(mh_file)
                    st.session_state.man_hours_df = mh_df
                    st.session_state.day_labels   = day_lbls
                    st.success(f"Man Hours: {len(mh_df)} employees loaded")
                except Exception as e:
                    errors.append(f"Man Hours: {e}")
            for file_obj, rtype in [(pick_file,"Picking"),(pack_file,"Packing"),(put_file,"Putaway"),(recv_file,"Receiving")]:
                if file_obj:
                    try:
                        df,dv,target,hours = load_zip_csv(file_obj, rtype)
                        if not df.empty:
                            rate_frames.append(df)
                            rate_meta[rtype] = {"date":dv,"target":target,"hours":hours}
                            st.success(f"{rtype}: {len(df)} users | Hours: {hours}")
                    except Exception as e:
                        errors.append(f"{rtype}: {e}")
            if rate_frames:
                st.session_state.rate_dfs  = rate_frames
                st.session_state.rate_meta = rate_meta
                all_rates = pd.concat(rate_frames, ignore_index=True)
                if st.session_state.man_hours_df is not None:
                    st.session_state.uph_df = compute_uph(all_rates, st.session_state.man_hours_df)
                    st.success(f"UPH: {len(st.session_state.uph_df)} matched workers")
            for e in errors: st.error(e)

    st.markdown("---")
    st.markdown("**Level colours:**")
    st.markdown("🟡 Trainee = up to 2 weeks")
    st.markdown("🟢 Starter = more than 1 month")
    st.markdown("🔵 Competent = 3 months+")
    st.markdown("🩷 Master = 6 months+")

tab1,tab2,tab3,tab4,tab5 = st.tabs(["Overview","User Detail","Daily Hours","Leaderboard","Days Rate Detail"])

# ── TAB 1: OVERVIEW ────────────────────────────────────────────────────────────
with tab1:
    mh = st.session_state.man_hours_df; day_labels = st.session_state.day_labels
    if mh is None:
        st.info("Upload files and click Process Files to get started.")
    else:
        # Level breakdown with colour-coded cards
        st.markdown('<div class="section-header">Workforce Level Breakdown</div>', unsafe_allow_html=True)
        lc = mh["Occupation Level"].value_counts()
        lv1,lv2,lv3,lv4,lv5 = st.columns(5)
        with lv1: st.markdown(f'<div class="metric-card"><h3>Total Employees</h3><p>{len(mh):,}</p></div>', unsafe_allow_html=True)
        with lv2: st.markdown(f'<div class="metric-card mc-yellow"><h3>🟡 Trainees</h3><p>{lc.get("Trainee",0)}</p></div>', unsafe_allow_html=True)
        with lv3: st.markdown(f'<div class="metric-card mc-teal"><h3>🟢 Starters</h3><p>{lc.get("Starter",0)}</p></div>', unsafe_allow_html=True)
        with lv4: st.markdown(f'<div class="metric-card mc-blue"><h3>🔵 Competent</h3><p>{lc.get("Competent",0)}</p></div>', unsafe_allow_html=True)
        with lv5: st.markdown(f'<div class="metric-card mc-pink"><h3>🩷 Masters</h3><p>{lc.get("Master",0)}</p></div>', unsafe_allow_html=True)

        occ_df = mh["Occupation Level"].value_counts().reset_index()
        occ_df.columns = ["Level","Count"]
        col_map = {"Trainee":"#ffeaa7","Starter":"#a8e6cf","Competent":"#74b9ff","Master":"#fd79a8"}
        st.altair_chart(alt.Chart(occ_df).mark_bar(cornerRadiusTopLeft=6,cornerRadiusTopRight=6).encode(
            x=alt.X("Level:N",sort=["Trainee","Starter","Competent","Master"]),y=alt.Y("Count:Q"),
            color=alt.Color("Level:N",scale=alt.Scale(domain=list(col_map.keys()),range=list(col_map.values())),legend=None),
            tooltip=["Level","Count"]
        ).properties(height=220), use_container_width=True)

        # Daily attendance
        st.markdown('<div class="section-header">Daily Attendance</div>', unsafe_allow_html=True)
        if day_labels:
            selected_day = st.selectbox("Select Day", day_labels)
            in_col = f"{selected_day}_In"; out_col = f"{selected_day}_Out"; hrs_col = f"{selected_day}_hrs"
            if in_col in mh.columns:
                present_mask = mh[in_col].notna() & (~mh[in_col].astype(str).str.strip().isin(["","nan","NaT"]))
                present_df = mh[present_mask].copy(); absent_df = mh[~present_mask].copy()
            else:
                present_df = pd.DataFrame(); absent_df = mh.copy()

            uph = st.session_state.uph_df
            a1,a2,a3,a4 = st.columns(4)
            total_hrs_day = present_df[hrs_col].sum() if hrs_col in present_df.columns and not present_df.empty else 0
            avg_uph_val   = uph["UPH"].mean() if not uph.empty else 0
            with a1: st.markdown(f'<div class="metric-card mc-green"><h3>Signed In</h3><p>{len(present_df)}</p></div>', unsafe_allow_html=True)
            with a2: st.markdown(f'<div class="metric-card mc-red"><h3>Off / Absent</h3><p>{len(absent_df)}</p></div>', unsafe_allow_html=True)
            with a3: st.markdown(f'<div class="metric-card"><h3>Hours Clocked</h3><p>{total_hrs_day:.1f} hrs</p></div>', unsafe_allow_html=True)
            with a4: st.markdown(f'<div class="metric-card mc-orange"><h3>Avg UPH</h3><p>{avg_uph_val:.1f}</p></div>', unsafe_allow_html=True)

            st.markdown(f'<div class="section-header">Signed In - {selected_day} ({len(present_df)} employees)</div>', unsafe_allow_html=True)
            if not present_df.empty:
                dept_ov = st.selectbox("Filter by Department",["All"]+sorted(mh["Department"].dropna().unique().tolist()),key="ov_dept")
                sp = present_df[["Cust-Oracle Username","Name Surname","Company Name","Department","Occupation","Occupation Level","Weeks On Site"]].copy()
                sp.columns = ["Username","Full Name","Company","Department","Occupation","Level","Weeks"]
                if in_col  in present_df.columns: sp["Sign In"]      = present_df[in_col].astype(str).values
                if out_col in present_df.columns: sp["Sign Out"]     = present_df[out_col].astype(str).values
                if hrs_col in present_df.columns: sp["Hours Worked"] = present_df[hrs_col].round(2).values
                if dept_ov != "All": sp = sp[sp["Department"] == dept_ov]
                st.dataframe(style_with_levels(sp.reset_index(drop=True)), use_container_width=True, hide_index=True)

                if hrs_col in present_df.columns:
                    dh = present_df.groupby("Department")[hrs_col].sum().reset_index()
                    dh.columns = ["Department","Hours"]
                    dh = dh[dh["Hours"]>0].sort_values("Hours",ascending=False)
                    if not dh.empty:
                        st.markdown('<div class="section-header">Hours by Department</div>', unsafe_allow_html=True)
                        st.altair_chart(alt.Chart(dh).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5,color="#2d6a9f").encode(
                            x=alt.X("Hours:Q"),y=alt.Y("Department:N",sort="-x"),tooltip=["Department","Hours"]
                        ).properties(height=max(200,len(dh)*28)),use_container_width=True)

            st.markdown(f'<div class="section-header">Off / Not Signed In - {selected_day} ({len(absent_df)} employees)</div>', unsafe_allow_html=True)
            if not absent_df.empty:
                sa = absent_df[["Cust-Oracle Username","Name Surname","Company Name","Department","Occupation","Occupation Level","Weeks On Site"]].copy()
                sa.columns = ["Username","Full Name","Company","Department","Occupation","Level","Weeks"]
                st.dataframe(style_with_levels(sa.reset_index(drop=True)),use_container_width=True,hide_index=True)

# ── TAB 2: USER DETAIL ─────────────────────────────────────────────────────────
with tab2:
    st.markdown('<div class="section-header">User Detail</div>', unsafe_allow_html=True)
    mh=st.session_state.man_hours_df; uph=st.session_state.uph_df; day_labels=st.session_state.day_labels
    if mh is None:
        st.info("Upload files and process them first.")
    else:
        fc1,fc2,fc3=st.columns(3)
        with fc1: dept_f=st.selectbox("Department",["All"]+sorted(mh["Department"].dropna().unique().tolist()))
        with fc2: comp_f=st.selectbox("Company",   ["All"]+sorted(mh["Company Name"].dropna().unique().tolist()))
        with fc3: lvl_f =st.selectbox("Level",     ["All","Trainee","Starter","Competent","Master"])
        dm=mh.copy()
        if dept_f!="All": dm=dm[dm["Department"]==dept_f]
        if comp_f!="All": dm=dm[dm["Company Name"]==comp_f]
        if lvl_f !="All": dm=dm[dm["Occupation Level"]==lvl_f]
        if not uph.empty:
            up2=uph.groupby("username").agg(UPH=("UPH","mean"),Report_Type=("Report Type",lambda x:", ".join(x.unique()))).reset_index()
            up2["UPH"]=up2["UPH"].round(1)
            dm=dm.merge(up2,left_on="Cust-Oracle Username",right_on="username",how="left")
        sc=["Cust-Oracle Username","Name Surname","Company Name","Department","Occupation","Occupation Level","Weeks On Site","Hire Date"]
        if "UPH" in dm.columns: sc+=["UPH","Report_Type"]
        search=st.text_input("Search by name or username")
        if search:
            dm=dm[dm["Name Surname"].astype(str).str.lower().str.contains(search.lower())|dm["Cust-Oracle Username"].astype(str).str.lower().str.contains(search.lower())]
        st.markdown(f"**{len(dm)} employees**")
        ft=dm[sc].rename(columns={"Cust-Oracle Username":"Username","Name Surname":"Full Name","Company Name":"Company","Occupation Level":"Level","Weeks On Site":"Weeks","Report_Type":"Role(s)"}).reset_index(drop=True)
        ft["Hire Date"]=pd.to_datetime(ft["Hire Date"],errors="coerce").dt.strftime("%Y-%m-%d")
        st.dataframe(style_with_levels(ft),use_container_width=True,hide_index=True)
        st.markdown("---")
        st.markdown('<div class="section-header">Individual User Profile</div>', unsafe_allow_html=True)
        uname=st.text_input("Enter Username:")
        if uname:
            um=mh[mh["Cust-Oracle Username"].astype(str).str.lower().str.strip()==uname.lower().strip()]
            if um.empty:
                st.warning(f"No employee found: {uname}")
            else:
                row=um.iloc[0]; level=row.get("Occupation Level","Trainee")
                st.markdown(f"""<div class="user-detail-card">
                  <h2 style="margin:0 0 8px;">{row.get("Name Surname","")}</h2>
                  <p><b>Username:</b> {row.get("Cust-Oracle Username","")}&nbsp;&nbsp;<b>Company:</b> {row.get("Company Name","")}&nbsp;&nbsp;<b>Department:</b> {row.get("Department","")}</p>
                  <p><b>Occupation:</b> {row.get("Occupation","")}&nbsp;&nbsp;<b>Hire Date:</b> {str(row.get("Hire Date",""))[:10]}&nbsp;&nbsp;<b>Service:</b> {row.get("Years of Service (Yrs/Mths)","")}</p>
                  <p><b>Weeks on Site:</b> {row.get("Weeks On Site",0)}&nbsp;&nbsp;<b>Level:</b> {badge_html(level)}</p>
                </div>""", unsafe_allow_html=True)
                hrs_data=[]; total_hrs=0
                for d in day_labels:
                    v=row.get(f"{d}_hrs",0); v=0 if pd.isna(v) else float(v)
                    hrs_data.append({"Day":d,"Hours":v}); total_hrs+=v
                hrs_df=pd.DataFrame(hrs_data)
                st.markdown('<div class="section-header">Hours Worked This Week</div>', unsafe_allow_html=True)
                st.altair_chart(alt.Chart(hrs_df).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5,color="#2d6a9f").encode(x="Day:N",y="Hours:Q",tooltip=["Day","Hours"]).properties(height=220),use_container_width=True)
                c1,c2=st.columns(2)
                c1.metric("Total Hours",f"{total_hrs:.1f} hrs")
                c2.metric("Avg Daily",f"{total_hrs/max(len(day_labels),1):.1f} hrs")
                if not uph.empty:
                    uu=uph[uph["username"].str.lower()==uname.lower()]
                    if not uu.empty:
                        st.markdown('<div class="section-header">Units Per Hour Performance</div>', unsafe_allow_html=True)
                        us=uu[["Report Type","Total Units","Active Hours","UPH","Report Date"]].copy()
                        us["Report Date"]=pd.to_datetime(us["Report Date"]).dt.strftime("%Y-%m-%d")
                        st.dataframe(us,use_container_width=True,hide_index=True)

# ── TAB 3: DAILY HOURS ─────────────────────────────────────────────────────────
with tab3:
    st.markdown('<div class="section-header">Daily and Weekly Hours Tracker</div>', unsafe_allow_html=True)
    mh=st.session_state.man_hours_df; day_labels=st.session_state.day_labels
    if mh is None:
        st.info("Upload files and process them first.")
    else:
        dept_sel=st.selectbox("Department",["All"]+sorted(mh["Department"].dropna().unique().tolist()),key="daily_dept")
        disp=mh[mh["Department"]==dept_sel].copy() if dept_sel!="All" else mh.copy()
        records=[]
        for _,r in disp.iterrows():
            rd={"Username":r.get("Cust-Oracle Username",""),"Name":r.get("Name Surname",""),"Company":r.get("Company Name",""),"Level":r.get("Occupation Level","")}
            wt=0
            for d in day_labels:
                v=r.get(f"{d}_hrs",0); v=0 if pd.isna(v) else float(v)
                rd[d]=round(v,2); wt+=v
            rd["Weekly Total"]=round(wt,2); records.append(rd)
        ht=pd.DataFrame(records)
        ht=ht[ht["Weekly Total"]>0].sort_values("Weekly Total",ascending=False)
        st.markdown(f"**{len(ht)} employees with recorded hours**")
        st.dataframe(style_with_levels(ht.reset_index(drop=True)),use_container_width=True,hide_index=True)
        if day_labels:
            st.markdown('<div class="section-header">Daily Total Hours Across Team</div>', unsafe_allow_html=True)
            dt_df=pd.DataFrame([{"Day":d,"Total Hours":ht[d].sum()} for d in day_labels if d in ht.columns])
            st.altair_chart(alt.Chart(dt_df).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5,color="#27ae60").encode(x="Day:N",y="Total Hours:Q",tooltip=["Day","Total Hours"]).properties(height=260),use_container_width=True)

# ── TAB 4: LEADERBOARD ─────────────────────────────────────────────────────────
with tab4:
    st.markdown('<div class="section-header">UPH Leaderboard - Occupation Matched</div>', unsafe_allow_html=True)
    uph=st.session_state.uph_df
    if uph.empty:
        st.info("Upload rate reports and process to see the leaderboard.")
    else:
        rts=st.selectbox("Report Type",["All"]+sorted(uph["Report Type"].unique().tolist()))
        du=uph[uph["Report Type"]==rts].copy() if rts!="All" else uph.copy()
        du=du.sort_values("UPH",ascending=False).reset_index(drop=True); du.index+=1
        sh=du[["Name Surname","username","Company","Department","Occupation","Occupation Level","Weeks On Site","Report Type","Total Units","Active Hours","UPH"]].copy()
        sh.columns=["Name","Username","Company","Department","Occupation","Level","Weeks","Type","Units","Active Hrs","UPH"]
        st.dataframe(style_with_levels(sh),use_container_width=True)
        top20=du.head(20)
        st.altair_chart(alt.Chart(top20).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5).encode(
            x=alt.X("UPH:Q"),y=alt.Y("Name Surname:N",sort="-x"),
            color=alt.Color("Occupation Level:N",scale=alt.Scale(domain=["Trainee","Starter","Competent","Master"],range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8"])),
            tooltip=["Name Surname","UPH","Occupation","Active Hours","Report Type"]
        ).properties(height=500,title="Top 20 - Units Per Hour"),use_container_width=True)

# ── TAB 5: DAYS RATE DETAIL ────────────────────────────────────────────────────
with tab5:
    st.markdown('<div class="section-header">Days Rate Detail - Hourly Breakdown Per User</div>', unsafe_allow_html=True)
    mh=st.session_state.man_hours_df; rate_dfs=st.session_state.rate_dfs; rate_meta=st.session_state.rate_meta
    if not rate_dfs:
        st.info("Upload at least one rate report (.zip) and click Process Files.")
    else:
        sel_type=st.selectbox("Report Type",list(rate_meta.keys()))
        rate_df=next((d for d in rate_dfs if d["report_type"].iloc[0]==sel_type),None)
        if rate_df is None:
            st.warning("No data for this report type.")
        else:
            meta=rate_meta.get(sel_type,{}); rpt_date=meta.get("date"); target_str=meta.get("target",""); hours_list=meta.get("hours",[])
            date_str=rpt_date.strftime("%d %B %Y") if rpt_date else "N/A"
            i1,i2,i3=st.columns(3)
            with i1: st.markdown(f'<div class="metric-card"><h3>Report Type</h3><p style="font-size:1.2rem">{sel_type}</p></div>',unsafe_allow_html=True)
            with i2: st.markdown(f'<div class="metric-card"><h3>Date</h3><p style="font-size:1.2rem">{date_str}</p></div>',unsafe_allow_html=True)
            with i3: st.markdown(f'<div class="metric-card"><h3>Target</h3><p style="font-size:.95rem">{target_str or "N/A"}</p></div>',unsafe_allow_html=True)

            hour_cols=[c for c in rate_df.columns if c.startswith("h_")]
            hour_display={hc:f"Hr {hc[2:]}" for hc in hour_cols}

            f1,f2,f3=st.columns(3)
            with f1: search_u=st.text_input("Search Username or Name",key="rs")
            with f2:
                d_opts=["All"]+(sorted(mh["Department"].dropna().unique().tolist()) if mh is not None else [])
                dept_r=st.selectbox("Department",d_opts,key="rd")
            with f3:
                c_opts=["All"]+(sorted(mh["Company Name"].dropna().unique().tolist()) if mh is not None else [])
                comp_r=st.selectbox("Company",c_opts,key="rc")

            records=[]
            for _,row in rate_df.iterrows():
                username=str(row["username"]).strip()
                full_name=username; company="-"; dept="-"; occupation="-"; hire_date="-"; weeks=0; level="-"; years_svc="-"
                if mh is not None:
                    mm=mh[mh["Cust-Oracle Username"].astype(str).str.strip().str.lower()==username.lower()]
                    if not mm.empty:
                        mhr=mm.iloc[0]
                        full_name=mhr.get("Name Surname",username); company=mhr.get("Company Name","-")
                        dept=mhr.get("Department","-"); occupation=mhr.get("Occupation","-")
                        hire_date=str(mhr.get("Hire Date",""))[:10]; weeks=mhr.get("Weeks On Site",0)
                        level=mhr.get("Occupation Level","Trainee"); years_svc=mhr.get("Years of Service (Yrs/Mths)","-")
                rec={"Username":username,"Full Name":full_name,"Company":company,"Department":dept,
                     "Occupation":occupation,"Level":level,"Weeks":weeks,"Hire Date":hire_date,
                     "Service":years_svc,"Total Units":int(row["Total"])}
                for hc in hour_cols:
                    val=row.get(hc,0)
                    rec[hour_display[hc]]=int(val) if not pd.isna(val) else 0
                active_hrs=int(row.get("_active_hrs",0))
                rec["Active Hrs"]=active_hrs
                rec["UPH"]=round(row["Total"]/active_hrs,1) if active_hrs>0 else 0
                records.append(rec)

            result_df=pd.DataFrame(records)
            if search_u: result_df=result_df[result_df["Username"].str.lower().str.contains(search_u.lower())|result_df["Full Name"].str.lower().str.contains(search_u.lower())]
            if dept_r!="All": result_df=result_df[result_df["Department"]==dept_r]
            if comp_r!="All": result_df=result_df[result_df["Company"]==comp_r]
            result_df=result_df.sort_values("Total Units",ascending=False).reset_index(drop=True)

            st.markdown('<div class="section-header">Summary</div>',unsafe_allow_html=True)
            k1,k2,k3,k4=st.columns(4)
            k1.metric("Workers Active",len(result_df))
            k2.metric("Total Units",f"{result_df['Total Units'].sum():,}")
            k3.metric("Avg UPH",f"{result_df['UPH'].mean():.1f}")
            k4.metric("Best UPH",f"{result_df['UPH'].max():.1f}")

            st.markdown(f'<div class="section-header">All Workers - {sel_type} | {date_str} | Hours shown: {", ".join(hours_list)}</div>',unsafe_allow_html=True)
            st.markdown(f"**{len(result_df)} workers**")
            st.dataframe(style_with_levels(result_df),use_container_width=True,hide_index=True)

            st.markdown("---")
            st.markdown('<div class="section-header">Individual Worker - Hourly Detail Report</div>',unsafe_allow_html=True)
            sel_user=st.selectbox("Select worker:",["-- Select --"]+result_df["Username"].tolist(),key="rsel")
            if sel_user!="-- Select --":
                ur=result_df[result_df["Username"]==sel_user].iloc[0]; lv=str(ur.get("Level","Trainee"))
                st.markdown(f"""<div class="user-detail-card">
                  <h2 style="margin:0 0 10px;">{ur["Full Name"]}</h2>
                  <div style="display:flex;flex-wrap:wrap;gap:20px;">
                    <div><b>Username:</b> {ur["Username"]}</div>
                    <div><b>Company:</b> {ur["Company"]}</div>
                    <div><b>Department:</b> {ur["Department"]}</div>
                    <div><b>Occupation:</b> {ur["Occupation"]}</div>
                    <div><b>Hire Date:</b> {ur["Hire Date"]}</div>
                    <div><b>Service:</b> {ur["Service"]}</div>
                    <div><b>Weeks:</b> {ur["Weeks"]}</div>
                    <div><b>Level:</b> {badge_html(lv)}</div>
                  </div>
                </div>""",unsafe_allow_html=True)

                hr_cols_display=[c for c in result_df.columns if c.startswith("Hr ")]
                hour_data=[{"Hour":hc,"Units":int(ur.get(hc,0))} for hc in hr_cols_display]
                hour_df=pd.DataFrame(hour_data); non_zero=hour_df[hour_df["Units"]>0]

                uh1,uh2,uh3,uh4=st.columns(4)
                uh1.metric("Total Units",f"{int(ur['Total Units']):,}")
                uh2.metric("Active Hours",ur["Active Hrs"])
                uh3.metric("UPH",ur["UPH"])
                peak_h=non_zero.loc[non_zero["Units"].idxmax(),"Hour"] if not non_zero.empty else "-"
                uh4.metric("Peak Hour",peak_h)

                st.markdown('<div class="section-header">Units Produced Per Hour</div>',unsafe_allow_html=True)
                if hour_df["Units"].sum()>0:
                    q66=hour_df["Units"].quantile(0.66); q33=hour_df["Units"].quantile(0.33)
                    bar=alt.Chart(hour_df).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5).encode(
                        x=alt.X("Hour:N",sort=None,title="Hour"),y=alt.Y("Units:Q",title="Units Produced"),
                        color=alt.condition(alt.datum.Units>=q66,alt.value("#27ae60"),
                               alt.condition(alt.datum.Units>=q33,alt.value("#f39c12"),alt.value("#e74c3c"))),
                        tooltip=["Hour","Units"]
                    ).properties(height=320,title=f"{ur['Full Name']} - Hourly Production ({sel_type})")
                    st.altair_chart(bar,use_container_width=True)

                st.markdown('<div class="section-header">Hourly Detail Table</div>',unsafe_allow_html=True)
                htbl=hour_df.copy()
                htbl["Status"]=htbl["Units"].apply(lambda v:"Active" if v>0 else "No Activity")
                st.dataframe(htbl,use_container_width=True,hide_index=True)

            st.markdown('<div class="section-header">All Workers - Units Per Hour Chart</div>',unsafe_allow_html=True)
            cd=result_df[result_df["UPH"]>0].sort_values("UPH",ascending=False).head(30)
            if not cd.empty:
                st.altair_chart(alt.Chart(cd).mark_bar(cornerRadiusTopLeft=5,cornerRadiusTopRight=5).encode(
                    x=alt.X("UPH:Q",title="Units Per Hour"),y=alt.Y("Full Name:N",sort="-x"),
                    color=alt.Color("Level:N",scale=alt.Scale(domain=["Trainee","Starter","Competent","Master","-"],range=["#ffeaa7","#a8e6cf","#74b9ff","#fd79a8","#dfe6e9"])),
                    tooltip=["Full Name","Username","Occupation","UPH","Total Units","Active Hrs","Level"]
                ).properties(height=600,title=f"Top 30 by UPH - {sel_type}"),use_container_width=True)
