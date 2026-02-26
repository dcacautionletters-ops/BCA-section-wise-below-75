import streamlit as st
import pandas as pd
import io
import time
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION & ADVANCED GLASS-MORPHISM CSS ---
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
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(15px);
        -webkit-backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        padding: 25px;
        margin: 10px 0;
        transition: transform 0.3s ease;
        text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    .glass-metric:hover {
        background: rgba(255, 255, 255, 0.08);
        transform: translateY(-5px);
        border: 1px solid rgba(146, 254, 157, 0.4);
    }
    .metric-title { color: #94a3b8; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; }
    .metric-value { font-size: 42px; font-weight: 800; color: #92fe9d; text-shadow: 0 0 15px rgba(146, 254, 157, 0.3); }
    
    .portal-status {
        font-family: 'Consolas', monospace; color: #92fe9d; font-size: 13px; padding: 15px;
        background: rgba(0, 0, 0, 0.3); border-radius: 4px; border-left: 4px solid #00d2ff; margin-bottom: 15px;
    }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #64748b; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SESSION STATE & AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if 'booted' not in st.session_state: st.session_state.booted = False

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        if not st.session_state.booted:
            st.markdown('<div class="portal-status">> Initiating Neural Link...<br>> Universal Engine Online.</div>', unsafe_allow_html=True)
            time.sleep(0.8); st.session_state.booted = True; st.rerun()
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. UNIVERSAL CORE LOGIC ---
BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

with st.sidebar:
    st.markdown("### 🛠️ Global Parameters")
    threshold = st.slider("Shortage Threshold (%)", 50, 95, 75, 5)
    st.divider()
    if st.button("System Logout"):
        st.session_state.authenticated = False
        st.rerun()

def apply_styles(ws, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
    warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), header_fill, border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = 20

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if not is_summary and cell.column > 4:
                try:
                    if cell.value != "" and cell.value is not None:
                        val = float(cell.value)
                        if val < threshold:
                            cell.fill, cell.font = crit_fill, Font(bold=True, color="FFFFFF")
                        else:
                            cell.fill, cell.font = warn_fill, Font(bold=True, color="000000")
                except: pass

def process_grid(data_df, cols, all_subjects):
    # Ensure Numeric Conversion to prevent TypeError
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce')
    
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    for sub in all_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None
        full_grid[sub] = pd.to_numeric(full_grid[sub], errors='coerce')

    sub_list = [s for s in all_subjects if s in full_grid.columns]
    theory_cols = [c for c in sub_list if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
    full_grid['Final Avg'] = full_grid[sub_list].mean(axis=1).round(2)
    
    mask = (full_grid[sub_list] < threshold).any(axis=1)
    shortage_grid = full_grid[mask].copy()
    
    if shortage_grid.empty: return None, None
    
    # Store subject-wise counts for analysis
    sub_shortage_counts = (shortage_grid[sub_list] < threshold).sum()
    
    for sub in sub_list:
        shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and x < threshold) else "")
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    final_cols = ['Sl No.', cols['roll'], cols['name'], cols['batch']] + sub_list + ['Theory Avg', 'Final Avg']
    return shortage_grid[final_cols], sub_shortage_counts

# --- 4. DASHBOARD INTERFACE ---
st.markdown('<p class="welcome-note">Institutional Precision Hub</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📂 Upload Universal Attendance File", type=["xlsx"])

if uploaded_file:
    raw_head = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in raw_head.iterrows():
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

    # Universal Dept Detection
    df['Dept_Prefix'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    all_depts = sorted(df['Dept_Prefix'].unique())
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Prevent IndexError by creating status sheet
        pd.DataFrame({"Status": ["System Clean"], "Threshold": [threshold]}).to_excel(writer, sheet_name="Audit_Log", index=False)
        
        master_summaries = []
        subject_impact_data = pd.Series(dtype=float)

        tabs = st.tabs(["📊 COMMAND CENTER"] + [f"💎 {d}" for d in all_depts])
        
        # Process Logic
        for d_idx, dept in enumerate(all_depts):
            d_df = df[df['Dept_Prefix'] == dept]
            core_subs = sorted([s for s in d_df[c_map['subject']].unique() if not any(b in str(s).upper() for b in BLACKLIST)])
            sections = sorted(d_df[c_map['batch']].unique())
            
            for sec in sections:
                sec_df = d_df[d_df[c_map['batch']] == sec]
                grid, sub_counts = process_grid(sec_df, c_map, core_subs)
                
                if grid is not None:
                    sh_name = str(sec).replace("/", "-")[:31]
                    grid.to_excel(writer, sheet_name=sh_name, index=False)
                    apply_styles(writer.sheets[sh_name])
                    master_summaries.append({'Section': sec, 'Count': len(grid), 'Dept': dept})
                    subject_impact_data = subject_impact_data.add(sub_counts, fill_value=0)
                    with tabs[d_idx+1]:
                        st.write(f"✅ {sec} processed: {len(grid)} shortages.")

        with tabs[0]:
            if master_summaries:
                summary_df = pd.DataFrame(master_summaries)
                
                # Metric Cards
                m_cols = st.columns(min(len(summary_df), 4))
                for idx, row in summary_df.iterrows():
                    with m_cols[idx % 4]:
                        st.markdown(f"""<div class="glass-metric">
                            <div class="metric-title">{row['Section']}</div>
                            <div class="metric-value">{row['Count']}</div>
                            <div style="color:#64748b; font-size:11px;">Students below {threshold}%</div>
                        </div>""", unsafe_allow_html=True)

                # Charts
                c1, c2 = st.columns(2)
                with c1:
                    fig1 = px.bar(summary_df, x='Section', y='Count', color='Dept', 
                                 title="Shortage by Section", template="plotly_dark",
                                 color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig1, use_container_width=True)
                
                with c2:
                    impact_df = subject_impact_data.reset_index()
                    impact_df.columns = ['Subject', 'Students']
                    fig2 = px.pie(impact_df.sort_values(by='Students', ascending=False).head(10), 
                                 names='Subject', values='Students', hole=0.4,
                                 title="Most Impacted Subjects (Top 10)", template="plotly_dark")
                    st.plotly_chart(fig2, use_container_width=True)

                summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                apply_styles(writer.sheets['MASTER SUMMARY'], is_summary=True)
            else:
                st.success(f"🌟 Institutional Excellence: No students below {threshold}% found.")

    st.download_button("📥 Extract Universal Magic Report", output.getvalue(), "VMS_Global_Attendance.xlsx", use_container_width=True)

st.markdown('<div class="footer">Universal VMS Reporting System v8.0 | Institutional Intelligence Active</div>', unsafe_allow_html=True)
