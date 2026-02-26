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
        padding: 25px; margin: 10px 0; text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    .metric-value { font-size: 42px; font-weight: 800; color: #92fe9d; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #64748b; font-size: 11px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC ---
BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def apply_styles(ws, threshold, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
    warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 20

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")
            if not is_summary and cell.column > 5:
                try:
                    val = float(cell.value)
                    if val < 70: cell.fill, cell.font = crit_fill, Font(bold=True, color="FFFFFF")
                    elif 70 <= val < threshold: cell.fill, cell.font = warn_fill, Font(bold=True, color="000000")
                except: pass

def process_grid(data_df, cols, batch_subjects, threshold):
    if data_df.empty: return None, None
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce')
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    final_subjects = [s for s in batch_subjects if str(s).upper() not in BLACKLIST]
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None
        full_grid[sub] = pd.to_numeric(full_grid[sub], errors='coerce')

    theory_cols = [c for c in final_subjects if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
    full_grid['Final Avg'] = full_grid[final_subjects].mean(axis=1).round(2)
    
    mask = (full_grid[final_subjects] < threshold).any(axis=1)
    shortage_grid = full_grid[mask].copy()
    if shortage_grid.empty: return None, None
    
    sub_counts = (shortage_grid[final_subjects] < threshold).sum()
    for sub in final_subjects:
        shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and x < threshold) else "")
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    final_cols = ['Sl No.', cols['roll'], cols['name'], cols['batch'], cols['sem']] + final_subjects + ['Theory Avg', 'Final Avg']
    return shortage_grid[final_cols], sub_counts

# --- 4. DASHBOARD INTERFACE ---
uploaded_file = st.file_uploader("📂 Upload Universal Attendance File", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in df_raw.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values): h_row = i; break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    c_map = {'sem': df.columns[5]} 
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    
    with st.sidebar:
        st.markdown("### 🛠️ Global Parameters")
        threshold = st.slider("Shortage Threshold (%)", 50, 95, 75, 5)
        available_depts = sorted(df['Dept'].unique())
        dept_choice = st.selectbox("Select Department", ["All Departments"] + available_depts)
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    if dept_choice != "All Departments":
        df = df[df['Dept'] == dept_choice]
        active_depts = [dept_choice]
    else:
        active_depts = available_depts

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({"Report": [f"Filtered: {dept_choice}"]}).to_excel(writer, sheet_name="Audit", index=False)
        summaries, subject_impact = [], pd.Series(dtype=float)
        
        tabs = st.tabs(["📊 COMMAND CENTER"] + [f"💎 {d}" for d in active_depts])

        for d_idx, dept in enumerate(active_depts):
            d_df = df[df['Dept'] == dept]
            series_list = sorted(d_df[c_map['batch']].astype(str).apply(lambda x: ' '.join(x.split()[:2])).unique())
            
            with tabs[d_idx+1]:
                for series in series_list:
                    s_df = d_df[d_df[c_map['batch']].astype(str).str.contains(series)]
                    s_subs = sorted([s for s in s_df[c_map['subject']].unique() if str(s).upper() not in BLACKLIST])
                    
                    gen_grid, _ = process_grid(s_df, c_map, s_subs, threshold)
                    if gen_grid is not None:
                        with st.expander(f"👁️ {series} GENERAL SUMMARY"):
                            st.dataframe(gen_grid, hide_index=True, use_container_width=True)
                        sn = f"{series} GEN"[:31]
                        gen_grid.to_excel(writer, sheet_name=sn, index=False)
                        apply_styles(writer.sheets[sn], threshold)
                    
                    sections = sorted(s_df[c_map['batch']].unique())
                    for sec in sections:
                        sec_df = s_df[s_df[c_map['batch']] == sec]
                        grid, counts = process_grid(sec_df, c_map, s_subs, threshold)
                        if grid is not None:
                            with st.expander(f"👁️ {sec}: {len(grid)} Shortages"):
                                st.dataframe(grid, hide_index=True, use_container_width=True)
                            sn_sec = str(sec).replace("/", "-")[:31]
                            grid.to_excel(writer, sheet_name=sn_sec, index=False)
                            apply_styles(writer.sheets[sn_sec], threshold)
                            summaries.append({'Section': sec, 'Count': len(grid)})
                            subject_impact = subject_impact.add(counts, fill_value=0)

        with tabs[0]:
            if summaries:
                sum_df = pd.DataFrame(summaries)
                m_cols = st.columns(min(len(sum_df), 4))
                for idx, row in sum_df.iterrows():
                    with m_cols[idx % 4]:
                        st.markdown(f'<div class="glass-metric"><div class="metric-title">{row["Section"]}</div><div class="metric-value">{row["Count"]}</div></div>', unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1: st.plotly_chart(px.bar(sum_df, x='Section', y='Count', color='Section', template="plotly_dark"), use_container_width=True)
                with c2:
                    # SAFETY FIX: Ensure index is reset properly and dataframe is not empty
                    if not subject_impact.empty and subject_impact.sum() > 0:
                        impact_df = subject_impact.reset_index()
                        impact_df.columns = ['Subject', 'Students']
                        # Final check for Plotly
                        if not impact_df.empty:
                            st.plotly_chart(px.pie(impact_df.head(10), names='Subject', values='Students', hole=0.4, title="Top Subject Impact", template="plotly_dark"), use_container_width=True)
                sum_df.to_excel(writer, sheet_name='SUMMARY', index=False)
            else:
                st.success(f"No shortages found for {dept_choice}.")

    st.download_button(f"📥 Download {dept_choice} Report", output.getvalue(), f"VMS_{dept_choice}_Report.xlsx", use_container_width=True)

st.markdown('<div class="footer">Universal VMS v10.1 | Safety Link Active</div>', unsafe_allow_html=True)
