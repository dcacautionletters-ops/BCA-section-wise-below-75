import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- UI CONFIG ---
st.set_page_config(page_title="Department of Computer Applications", layout="wide")

# --- CUSTOM BRANDING, ANIMATIONS & BLINKING EYE ---
st.markdown("""
    <style>
    @keyframes fadeIn {
        0% { opacity: 0; transform: translateY(10px); }
        100% { opacity: 1; transform: translateY(0); }
    }
    .main-title { font-size:32px !important; font-weight: bold; color: #2C3E50; margin-bottom: 5px; }
    .welcome-note { 
        font-size: 52px !important; 
        font-weight: 800;
        color: #1F4E78; 
        margin-bottom: 0px;
        animation: fadeIn 2.5s ease-in-out;
    }
    .magic-text { 
        font-size: 20px !important; 
        color: #7F8C8D; 
        margin-bottom: 10px;
        font-style: italic;
        animation: fadeIn 3.5s ease-in-out;
    }
    .magician-icon { font-size: 80px !important; margin-top: 10px; animation: fadeIn 4s ease-in-out; }
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #BDC3C7; font-size: 14px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="main-title">🏛️ Department of Computer Applications</p>', unsafe_allow_html=True)
st.markdown('<p class="welcome-note">Welcome to Presidency.</p>', unsafe_allow_html=True)
st.markdown('<p class="magic-text">You are about to experience a Magic for reporting here.</p>', unsafe_allow_html=True)
st.markdown('<p class="magician-icon">🧙‍♂️</p>', unsafe_allow_html=True)

if 'magic_unlocked' not in st.session_state:
    st.session_state['magic_unlocked'] = None

st.write("### Ready to reveal the magic?")
col1, col2, _ = st.columns([1, 1, 5])
with col1:
    if st.button("✅ Yes"): st.session_state['magic_unlocked'] = True
with col2:
    if st.button("❌ No"): st.session_state['magic_unlocked'] = False

if st.session_state['magic_unlocked'] is True:
    BCA_BATCHES = ["BCA 2025", "BCA 2024", "BCA 2023"]
    MCA_BATCHES = ["MCA 2025", "MCA 2024"]
    THRESHOLDS = [75]
    BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]
    ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

    def apply_styles(ws, is_summary=False):
        thin_side = Side(style='thin', color="4D4D4D")
        full_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        total_header_fill = PatternFill(start_color="7F8C8D", end_color="7F8C8D", fill_type="solid")
        crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
        warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
        white_font = Font(bold=True, color="FFFFFF")
        black_font = Font(bold=True, color="000000")

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = white_font
            cell.fill = total_header_fill if not is_summary and col >= (ws.max_column - 2) else header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = full_border
            ws.column_dimensions[cell.column_letter].width = 20

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = full_border
                cell.alignment = Alignment(horizontal="center")
                if not is_summary and 4 < cell.column < (ws.max_column - 2) and cell.value is not None:
                    try:
                        val = float(cell.value)
                        if val < 70:
                            cell.fill = crit_fill; cell.font = white_font
                        else:
                            cell.fill = warn_fill; cell.font = black_font
                    except: pass

    def create_excel_sheet(data_df, writer, sheet_name, cols):
        if data_df.empty: return
        grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
        grid.insert(0, 'Sl No.', range(1, len(grid) + 1))
        sub_cols = grid.columns[4:]
        grid['Theory Avg'] = grid[[c for c in sub_cols if "LAB" not in str(c).upper()]].mean(axis=1).round(2)
        grid['Final Avg'] = grid[sub_cols].mean(axis=1).round(2)
        
        final_sheet_name = sheet_name[:31]
        grid.to_excel(writer, sheet_name=final_sheet_name, index=False)
        apply_styles(writer.sheets[final_sheet_name])
        return grid

    def run_batch_logic(df, batches, writer, cols):
        batch_summaries = []
        for group in batches:
            course, year = group.split()
            mask = (df[cols['batch']].astype(str).str.contains(course, case=False, na=False)) & \
                   (df[cols['batch']].astype(str).str.contains(year, na=False))
            batch_df = df[mask].copy()
            batch_df = batch_df[~batch_df[cols['subject']].astype(str).str.contains('|'.join(BLACKLIST), case=False, na=False)]
            
            if batch_df.empty: continue

            # 1. CREATE GENERAL BATCH REPORT (FULL)
            limit = THRESHOLDS[0]
            gen_shortage = batch_df[batch_df[cols['attendance']] < limit].copy()
            if not gen_shortage.empty:
                create_excel_sheet(gen_shortage, writer, f"{group} GENERAL", cols)
                with st.expander(f"👁️ {group} - FULL BATCH (Below {limit}%)"):
                    st.info(f"Showing all sections for {group}")
                    # No need to display huge dataframe here if sections are shown below, 
                    # but kept for consistency
                
            # 2. CREATE SECTION-WISE REPORTS
            batch_df['Section_Extracted'] = batch_df[cols['batch']].astype(str).apply(
                lambda x: x.split('-')[-1].strip() if '-' in x else "Gen"
            )
            sections = sorted(batch_df['Section_Extracted'].unique())
            
            for section in sections:
                sec_df = batch_df[batch_df['Section_Extracted'] == section]
                sec_shortage = sec_df[sec_df[cols['attendance']] < limit].copy()
                
                if not sec_shortage.empty:
                    create_excel_sheet(sec_shortage, writer, f"{group} Sec {section}", cols)
                    with st.expander(f"👁️ {group} - Section {section} (Below {limit}%)"):
                        st.dataframe(sec_shortage.pivot_table(index=[cols['roll'], cols['name']], 
                                     columns=cols['subject'], values=cols['attendance']).reset_index(), 
                                     hide_index=True, use_container_width=True)
                    
                    batch_summaries.append({'Batch': group, 'Section': section, 'Count': len(sec_shortage[cols['roll']].unique())})
        return batch_summaries

    st.divider()
    uploaded_file = st.file_uploader("📂 Upload Attendance Excel", type=["xlsx"])

    if uploaded_file:
        raw_df = pd.read_excel(uploaded_file, header=None).head(10)
        header_row = 0
        for i, row in raw_df.iterrows():
            row_str = " ".join(row.astype(str).values)
            if "Roll No" in row_str or "Batch" in row_str:
                header_row = i
                break
        
        df = pd.read_excel(uploaded_file, header=header_row)
        
        col_map = {}
        for c in df.columns:
            c_str = str(c).strip()
            if "Roll No" in c_str: col_map['roll'] = c
            elif "Student Name" in c_str: col_map['name'] = c
            elif "Batch" in c_str: col_map['batch'] = c
            elif "Course Name" in c_str or "Subject" in c_str: col_map['subject'] = c
            elif ATT_COL_NAME in c_str: col_map['attendance'] = c

        df[col_map['attendance']] = pd.to_numeric(df[col_map['attendance']], errors='coerce')
        output = io.BytesIO()
        tab1, tab2, tab3 = st.tabs(["🎓 BCA Analytics", "📜 MCA Analytics", "📊 Summary"])

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with tab1: bca_data = run_batch_logic(df, BCA_BATCHES, writer, col_map)
            with tab2: mca_data = run_batch_logic(df, MCA_BATCHES, writer, col_map)
            with tab3:
                summary_df = pd.DataFrame(bca_data + mca_data)
                st.subheader("📈 Section-wise Shortage (<75%)")
                if not summary_df.empty:
                    fig = px.bar(summary_df, x='Batch', y='Count', color='Section', barmode='group', text_auto=True)
                    st.plotly_chart(fig, use_container_width=True)
                    summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                    apply_styles(writer.sheets['MASTER SUMMARY'], is_summary=True)

        st.download_button(label="📥 Download Magic Report", data=output.getvalue(), file_name="DCA_Sectionwise_Attendance_Report.xlsx")

st.markdown('<div class="footer">© VMS</div>', unsafe_allow_html=True)
