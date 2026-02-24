import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import re

# --- UI CONFIG ---
st.set_page_config(page_title="Department of Computer Applications", layout="wide")

# --- CSS STYLING ---
st.markdown("""
    <style>
    @keyframes fadeIn { 0% { opacity: 0; transform: translateY(10px); } 100% { opacity: 1; transform: translateY(0); } }
    .main-title { font-size:32px !important; font-weight: bold; color: #2C3E50; margin-bottom: 5px; }
    .welcome-note { font-size: 52px !important; font-weight: 800; color: #1F4E78; margin-bottom: 0px; animation: fadeIn 2.5s ease-in-out; }
    .magic-text { font-size: 20px !important; color: #7F8C8D; margin-bottom: 10px; font-style: italic; animation: fadeIn 3.5s ease-in-out; }
    .magician-icon { font-size: 80px !important; margin-top: 10px; animation: fadeIn 4s ease-in-out; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #BDC3C7; font-size: 14px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="main-title">🏛️ Department of Computer Applications</p>', unsafe_allow_html=True)
st.markdown('<p class="welcome-note">Welcome to Presidency.</p>', unsafe_allow_html=True)
st.markdown('<p class="magic-text">Precision Reporting: General + Section-Wise.</p>', unsafe_allow_html=True)
st.markdown('<p class="magician-icon">🧙‍♂️</p>', unsafe_allow_html=True)

if 'magic_unlocked' not in st.session_state:
    st.session_state['magic_unlocked'] = None

st.write("### Ready to reveal the magic?")
c1, c2, _ = st.columns([1, 1, 5])
with c1: 
    if st.button("✅ Yes"): st.session_state['magic_unlocked'] = True
with c2: 
    if st.button("❌ No"): st.session_state['magic_unlocked'] = False

if st.session_state['magic_unlocked'] is True:
    BCA_BATCHES = ["BCA 2025", "BCA 2024", "BCA 2023"]
    MCA_BATCHES = ["MCA 2025", "MCA 2024"]
    BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
    ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

    def apply_styles(ws, is_summary=False):
        thin = Side(style='thin', color="4D4D4D")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
        warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
        white_font = Font(bold=True, color="FFFFFF")
        
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = white_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[cell.column_letter].width = 20

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
                if not is_summary and 4 < cell.column < (ws.max_column - 2) and cell.value:
                    try:
                        val = float(cell.value)
                        if val < 70:
                            cell.fill = crit_fill; cell.font = white_font
                        else:
                            cell.fill = warn_fill; cell.font = Font(bold=True, color="000000")
                    except: pass

    def extract_section(batch_str):
        match = re.search(r'[- ]([A-Z])$', str(batch_str).strip())
        return match.group(1) if match else "Gen"

    def process_grid(data_df, cols, all_subjects):
        if data_df.empty: return None
        
        # Pivot everything first to calculate averages accurately
        full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                        columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
        
        # Ensure all core subjects exist in the columns
        for sub in all_subjects:
            if sub not in full_grid.columns:
                full_grid[sub] = None

        # Reorder columns: Sl No, Roll, Name, Batch, [Subjects], Averages
        sub_list = [s for s in all_subjects if s in full_grid.columns]
        
        # Averages Calculation
        theory_cols = [c for c in sub_list if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
        
        full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
        full_grid['Final Avg'] = full_grid[sub_list].mean(axis=1).round(2)
        
        # Filter: Only keep students who have at least one subject < 75%
        mask = (full_grid[sub_list] < 75).any(axis=1)
        shortage_grid = full_grid[mask].copy()
        
        if shortage_grid.empty: return None

        # Masking: Only show values in subject columns if they are < 75
        for sub in sub_list:
            shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and x < 75) else "")

        shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
        
        # Final Column Order
        final_cols = ['Sl No.', cols['roll'], cols['name'], cols['batch']] + sub_list + ['Theory Avg', 'Final Avg']
        return shortage_grid[final_cols]

    def run_batch_logic(df, batches, writer, cols):
        batch_summaries = []
        for group in batches:
            course, year = group.split()
            mask = (df[cols['batch']].astype(str).str.contains(course, case=False, na=False)) & \
                   (df[cols['batch']].astype(str).str.contains(year, na=False))
            batch_df = df[mask].copy()
            
            # Identify core subjects (Excluding Blacklist)
            core_subjects = sorted([s for s in batch_df[cols['subject']].unique() 
                                   if not any(b in str(s).upper() for b in BLACKLIST)])
            
            if batch_df.empty: continue

            # --- GENERAL REPORT ---
            gen_grid = process_grid(batch_df, cols, core_subjects)
            if gen_grid is not None:
                sh_name = f"{group} ALL"[:31]
                gen_grid.to_excel(writer, sheet_name=sh_name, index=False)
                apply_styles(writer.sheets[sh_name])
                with st.expander(f"👁️ {group} - GENERAL (Shortage Count: {len(gen_grid)})"):
                    st.dataframe(gen_grid, hide_index=True)

            # --- SECTION-WISE ---
            batch_df['Ext_Sec'] = batch_df[cols['batch']].apply(extract_section)
            for sec in sorted(batch_df['Ext_Sec'].unique()):
                sec_df = batch_df[batch_df['Ext_Sec'] == sec]
                sec_grid = process_grid(sec_df, cols, core_subjects)
                if sec_grid is not None:
                    sh_name = f"{group} {sec}"[:31]
                    sec_grid.to_excel(writer, sheet_name=sh_name, index=False)
                    apply_styles(writer.sheets[sh_name])
                    batch_summaries.append({'Batch': group, 'Section': sec, 'Count': len(sec_grid)})
        
        return batch_summaries

    st.divider()
    uploaded_file = st.file_uploader("📂 Upload Attendance Excel", type=["xlsx"])

    if uploaded_file:
        raw_head = pd.read_excel(uploaded_file, header=None).head(10)
        h_row = 0
        for i, row in raw_head.iterrows():
            if any(k in str(row.values) for k in ["Roll No", "Batch"]):
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
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            t1, t2, t3 = st.tabs(["🎓 BCA", "📜 MCA", "📊 Summary"])
            with t1: bca_data = run_batch_logic(df, BCA_BATCHES, writer, c_map)
            with t2: mca_data = run_batch_logic(df, MCA_BATCHES, writer, c_map)
            with t3:
                summary_df = pd.DataFrame(bca_data + mca_data)
                if not summary_df.empty:
                    st.plotly_chart(px.bar(summary_df, x='Batch', y='Count', color='Section', barmode='group'))
                    summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                    apply_styles(writer.sheets['MASTER SUMMARY'], is_summary=True)

        st.download_button("📥 Download Final Report", output.getvalue(), "Attendance_Report_Fixed.xlsx")

st.markdown('<div class="footer">© VMS</div>', unsafe_allow_html=True)
