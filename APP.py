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
        padding: 20px; margin: 10px 0; text-align: center; border-left: 5px solid #92fe9d;
    }
    .search-box {
        background: rgba(255, 255, 255, 0.05); padding: 20px; border-radius: 15px;
        border: 1px solid rgba(0, 210, 255, 0.3); margin-bottom: 20px;
    }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #64748b; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">Universal Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC ---
BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def apply_styles(ws, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), header_fill, border
        ws.column_dimensions[cell.column_letter].width = 20
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

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

# --- 4. DATA PROCESSING & UI ---
st.markdown('<p class="welcome-note">Institutional Analytics Hub</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📂 Upload Attendance Excel", type=["xlsx"])

if uploaded_file:
    # Header Detection
    df_raw = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in df_raw.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i; break
    
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
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # We need to process the data first to populate Search and Charts
        master_summaries = []
        all_pivot_grids = []
        all_depts = sorted(df['Dept_Key'].unique())

        # Logic Execution
        for dept in all_depts:
            dept_df = df[df['Dept_Key'] == dept]
            core_subs = sorted([s for s in dept_df[c_map['subject']].unique() if not any(b in str(s).upper() for b in BLACKLIST)])
            sections = sorted(dept_df[c_map['batch']].unique())
            
            for sec in sections:
                sec_df = dept_df[dept_df[c_map['batch']] == sec]
                grid, sub_list = get_pivot_data(sec_df, c_map, core_subs)
                all_pivot_grids.append(grid)
                
                # Excel Reporting (Shortage Only)
                mask = (grid[sub_list] < 75).any(axis=1)
                shortage_grid = grid[mask].copy()
                if not shortage_grid.empty:
                    sh_name = str(sec)[:31]
                    shortage_grid.to_excel(writer, sheet_name=sh_name, index=False)
                    apply_styles(writer.sheets[sh_name])
                    master_summaries.append({'Section': sec, 'Shortage': len(shortage_grid), 'Dept': dept})

        # --- TAB RENDERING ---
        t_search, t_depts, t_stats = st.tabs(["🔍 GLOBAL SEARCH", "💎 DEPARTMENTAL VIEW", "📊 COMMAND CENTER"])

        with t_search:
            st.markdown('<div class="search-box">', unsafe_allow_html=True)
            search_query = st.text_input("🕵️ Search by Name or Roll Number...", placeholder="Type here...")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if search_query:
                combined_df = pd.concat(all_pivot_grids, ignore_index=True)
                results = combined_df[(combined_df[c_map['name']].astype(str).str.contains(search_query, case=False)) | 
                                      (combined_df[c_map['roll']].astype(str).str.contains(search_query, case=False))]
                if not results.empty:
                    st.dataframe(results, use_container_width=True, hide_index=True)
                else:
                    st.warning("No records found.")

        with t_depts:
            for dept in all_depts:
                with st.expander(f"📁 {dept} DEPARTMENT SECTIONS"):
                    dept_sum = [s for s in master_summaries if s['Dept'] == dept]
                    for s in dept_sum:
                        st.write(f"🔹 **{s['Section']}**: {s['Shortage']} students in shortage")

        with t_stats:
            summary_df = pd.DataFrame(master_summaries)
            if not summary_df.empty:
                # Colorful Bar Chart - Each Dept gets a color
                fig = px.bar(summary_df, x='Section', y='Shortage', color='Dept',
                           title="Multi-Department Shortage Analysis",
                           color_discrete_sequence=px.colors.qualitative.Bold,
                           template="plotly_dark")
                st.plotly_chart(fig, use_container_width=True)
                
                summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                apply_styles(writer.sheets['MASTER SUMMARY'])

    st.download_button("📥 Extract Universal Magic Report", output.getvalue(), "Universal_VMS_Report.xlsx")

st.markdown('<div class="footer">Designed by © VMS | Secure Universal Session Active</div>', unsafe_allow_html=True)
