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
        background: linear-gradient(rgba(15, 23, 42, 0.94), rgba(15, 23, 42, 0.94)), 
                    url("https://images.unsplash.com/photo-1451187580459-43490279c0fa?ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80");
        background-size: cover; background-attachment: fixed;
    }
    .welcome-note { 
        background: linear-gradient(to right, #00d2ff, #92fe9d); 
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
        font-size: 40px !important; font-weight: 800; text-align: center; margin-bottom: 20px;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px;
        padding: 15px; text-align: center; border-bottom: 3px solid #00d2ff;
    }
    .metric-value { font-size: 32px; font-weight: 800; color: #92fe9d; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #64748b; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">Institutional Reporting Link</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        p = st.text_input("Access Key", type="password")
        if st.button("Unlock Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Invalid Key")
    st.stop()

# --- 3. RECTIFIED LOGIC ---
BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def apply_styles(ws):
    thin = Side(style='thin', color="444444")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 22
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")

def get_pivot_data(df, cols, subjects):
    # Step 1: Force numeric conversion BEFORE pivoting to solve the TypeError
    df[cols['attendance']] = pd.to_numeric(df[cols['attendance']], errors='coerce')
    
    grid = df.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                          columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    for s in subjects:
        if s not in grid.columns: grid[s] = None
    
    # Step 2: Ensure pivoted columns are also numeric
    for col in subjects:
        grid[col] = pd.to_numeric(grid[col], errors='coerce')
        
    theory = [c for c in subjects if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    
    # Calculation with safety rounding
    grid['Theory Avg'] = grid[theory].mean(axis=1, skipna=True).astype(float).round(2)
    grid['Final Avg'] = grid[subjects].mean(axis=1, skipna=True).astype(float).round(2)
    return grid

# --- 4. DASHBOARD ENGINE ---
st.markdown('<p class="welcome-note">Multi-Department Analytics Hub</p>', unsafe_allow_html=True)
file = st.file_uploader("📂 Upload Attendance Excel", type=["xlsx"])

if file:
    # Header Detection
    raw = pd.read_excel(file, header=None).head(15)
    h_idx = 0
    for i, row in raw.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values): h_idx = i; break
    
    df = pd.read_excel(file, header=h_idx)
    c_map = {}
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    
    output = io.BytesIO()
    all_grids = []
    summaries = []

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Prevent IndexError by creating status sheet immediately
        pd.DataFrame({"Status": ["System Ready"], "Timestamp": [time.ctime()]}).to_excel(writer, sheet_name="Audit_Log", index=False)
        
        depts = sorted(df['Dept'].unique())
        for d in depts:
            d_df = df[df['Dept'] == d]
            subs = sorted([s for s in d_df[c_map['subject']].unique() if not any(b in str(s).upper() for b in BLACKLIST)])
            sections = sorted(d_df[c_map['batch']].unique())
            
            for sec in sections:
                s_df = d_df[d_df[c_map['batch']] == sec]
                grid = get_pivot_data(s_df, c_map, subs)
                all_grids.append(grid)
                
                mask = (grid[subs] < 75).any(axis=1)
                shortage = grid[mask].copy()
                if not shortage.empty:
                    for s in subs: shortage[s] = shortage[s].apply(lambda x: x if (pd.notnull(x) and x != "" and float(x or 0) < 75) else "")
                    shortage.insert(0, 'Sl No.', range(1, len(shortage) + 1))
                    sn = str(sec)[:31]
                    shortage.to_excel(writer, sheet_name=sn, index=False)
                    apply_styles(writer.sheets[sn])
                    summaries.append({'Section': sec, 'Count': len(shortage), 'Dept': d})

        if summaries:
            pd.DataFrame(summaries).to_excel(writer, sheet_name="MASTER_SUMMARY", index=False)
            apply_styles(writer.sheets["MASTER_SUMMARY"])

    # --- UI DISPLAY ---
    tab_search, tab_visuals = st.tabs(["🔍 GLOBAL SEARCH HUB", "📊 PERFORMANCE STATS"])

    with tab_search:
        query = st.text_input("🕵️ Find Student (Name or Roll No)")
        if query and all_grids:
            master_df = pd.concat(all_grids, ignore_index=True)
            res = master_df[(master_df[c_map['name']].astype(str).str.contains(query, case=False)) | 
                            (master_df[c_map['roll']].astype(str).str.contains(query, case=False))]
            st.dataframe(res, use_container_width=True, hide_index=True)

    with tab_visuals:
        if summaries:
            sum_df = pd.DataFrame(summaries)
            # Dynamic Vibrant Colors
            fig = px.bar(sum_df, x='Section', y='Count', color='Dept',
                         title="Shortage Distribution", 
                         color_discrete_sequence=px.colors.qualitative.Alphabet,
                         template="plotly_dark")
            st.plotly_chart(fig, use_container_width=True)
            
            cols = st.columns(len(depts))
            for i, d in enumerate(depts):
                val = sum_df[sum_df['Dept'] == d]['Count'].sum()
                with cols[i]:
                    st.markdown(f'<div class="glass-metric"><small>{d}</small><br><span class="metric-value">{val}</span></div>', unsafe_allow_html=True)
        else:
            st.success("✅ No shortages found.")

    st.download_button("📥 Download Excel Report", output.getvalue(), "VMS_Institutional_Report.xlsx", use_container_width=True)

st.markdown('<div class="footer">Universal VMS Neural Link v6.0 | Error-Free Mode Active</div>', unsafe_allow_html=True)
