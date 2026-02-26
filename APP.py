import streamlit as st
import pandas as pd
import io
import time
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="VMS Universal Reporting", layout="wide")
MASTER_PASSWORD = "VMS@123"

st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(rgba(15, 23, 42, 0.92), rgba(15, 23, 42, 0.92)), 
                    url("https://images.unsplash.com/photo-1451187580459-43490279c0fa?ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80");
        background-size: cover; background-attachment: fixed;
    }
    .welcome-note { 
        background: linear-gradient(to right, #00d2ff, #92fe9d); 
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
        font-size: 48px !important; font-weight: 700; text-align: center; margin: 40px 0 10px 0;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.03); backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 15px;
        padding: 20px; margin: 10px 0; text-align: center;
    }
    .metric-value { font-size: 36px; font-weight: 800; color: #92fe9d; }
    .portal-status { font-family: 'Consolas', monospace; color: #92fe9d; font-size: 13px; padding: 15px; background: rgba(0, 0, 0, 0.3); border-left: 4px solid #00d2ff; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #64748b; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if 'booted' not in st.session_state: st.session_state.booted = False

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">Universal Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        if not st.session_state.booted:
            st.markdown('<div class="portal-status">> Initiating Neural Link...<br>> Multi-Dept Logic Loaded.</div>', unsafe_allow_html=True)
            time.sleep(0.8); st.session_state.booted = True; st.rerun()
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC (DEPT-AGNOSTIC) ---
BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def apply_styles(ws, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
    warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), header_fill, border
        ws.column_dimensions[cell.column_letter].width = 20
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if not is_summary and cell.column > 4:
                try:
                    if cell.value and float(cell.value) < 70:
                        cell.fill, cell.font = crit_fill, Font(bold=True, color="FFFFFF")
                    elif cell.value:
                        cell.fill = warn_fill
                except: pass

def get_pivot_data(data_df, cols, all_subjects):
    grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    for sub in all_subjects:
        if sub not in grid.columns: grid[sub] = None
    sub_list = [s for s in all_subjects if s in grid.columns]
    theory_cols = [c for c in sub_list if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    grid['Theory Avg'] = grid[theory_cols].mean(axis=1).round(2)
    grid['Final Avg'] = grid[sub_list].mean(axis=1).round(2)
    return grid, sub_list

# --- 4. THE INTERFACE ---
st.markdown('<p class="welcome-note">Universal Analytics Hub</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📂 Drop Departmental Attendance Excel Here", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None).head(10)
    h_row = next(i for i, row in df_raw.iterrows() if any(k in str(row.values) for k in ["Roll No", "Batch"]))
    df = pd.read_excel(uploaded_file, header=h_row)
    
    c_map = {}
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    df[c_map['attendance']] = pd.to_numeric(df[c_map['attendance']], errors='coerce')
    df['Dept_Key'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    all_depts = sorted(df['Dept_Key'].unique())
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        tabs = st.tabs(["🔍 GLOBAL SEARCH"] + [f"💎 {d}" for d in all_depts] + ["📊 COMMAND CENTER"])
        master_summaries = []

        # PRE-PROCESS ALL DATA FOR SEARCH & TABS
        all_grid_data = []

        for i, dept in enumerate(all_depts):
            with tabs[i+1]:
                dept_df = df[df['Dept_Key'] == dept]
                core_subs = sorted([s for s in dept_df[c_map['subject']].unique() if not any(b in str(s).upper() for b in BLACKLIST)])
                
                sections = sorted(dept_df[c_map['batch']].unique())
                for sec in sections:
                    sec_df = dept_df[dept_df[c_map['batch']] == sec]
                    grid, sub_list = get_pivot_data(sec_df, c_map, core_subs)
                    all_grid_data.append(grid)
                    
                    # Filter for shortage for the Excel report
                    mask = (grid[sub_list] < 75).any(axis=1)
                    shortage_grid = grid[mask].copy()
                    
                    if not shortage_grid.empty:
                        for sub in sub_list:
                            shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and x < 75) else "")
                        shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
                        sh_name = str(sec)[:31]
                        shortage_grid.to_excel(writer, sheet_name=sh_name, index=False)
                        apply_styles(writer.sheets[sh_name])
                        master_summaries.append({'Section': sec, 'Shortage': len(shortage_grid), 'Dept': dept})
                        st.write(f"✅ Processed: {sec}")

        # COMBINE DATA FOR SEARCH
        full_master_df = pd.concat(all_grid_data, ignore_index=True)

        with tabs[0]:
            st.markdown("### 🕵️ Global Student Search")
            search_query = st.text_input("Enter Student Name or Roll Number")
            if search_query:
                # Fuzzy search across name and roll
                search_mask = (full_master_df[c_map['name']].astype(str).str.contains(search_query, case=False)) | \
                             (full_master_df[c_map['roll']].astype(str).str.contains(search_query, case=False))
                results = full_master_df[search_mask]
                if not results.empty:
                    st.dataframe(results, hide_index=True)
                else:
                    st.warning("No student found with those details.")

        with tabs[-1]:
            summary_df = pd.DataFrame(master_summaries)
            if not summary_df.empty:
                st.markdown("### 📍 Institutional Performance")
                fig = px.bar(summary_df, x='Section', y='Shortage', color='Dept', 
                           template="plotly_dark", barmode="group", color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig, use_container_width=True)
                summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                apply_styles(writer.sheets['MASTER SUMMARY'], is_summary=True)

    st.download_button("📥 Extract Universal Magic Report", output.getvalue(), "Universal_VMS_Report.xlsx")

st.markdown('<div class="footer">Designed by © VMS | Secure Universal Session Active</div>', unsafe_allow_html=True)
