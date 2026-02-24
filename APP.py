import streamlit as st
import pandas as pd
import io
import plotly.express as px # Added for enhanced magic charts
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
    
    @keyframes blink {
        0% { opacity: 1; }
        50% { opacity: 0.3; }
        100% { opacity: 1; }
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

    .magician-icon {
        font-size: 80px !important;
        margin-top: 10px;
        animation: fadeIn 4s ease-in-out;
    }

    .blinking-eye {
        animation: blink 2s infinite;
        display: inline-block;
    }
    
    .footer { position: fixed; bottom: 10px; width: 100%; text-align: center; color: #BDC3C7; font-size: 14px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- HEADER SECTION ---
st.markdown('<p class="main-title">🏛️ Department of Computer Applications</p>', unsafe_allow_html=True)
st.markdown('<p class="welcome-note">Welcome to Presidency.</p>', unsafe_allow_html=True)
st.markdown('<p class="magic-text">You are about to experience a Magic for reporting here.</p>', unsafe_allow_html=True)
st.markdown('<p class="magician-icon">🧙‍♂️</p>', unsafe_allow_html=True)

# --- INTERACTIVE MAGIC GATE ---
st.write("### Ready to reveal the magic?")
col1, col2, _ = st.columns([1, 1, 5])

if 'magic_unlocked' not in st.session_state:
    st.session_state['magic_unlocked'] = None

with col1:
    if st.button("✅ Yes"):
        st.session_state['magic_unlocked'] = True
with col2:
    if st.button("❌ No"):
        st.session_state['magic_unlocked'] = False

# Feedback logic
if st.session_state['magic_unlocked'] is True:
    st.balloons()
    st.success("### 💥 Boom! Congratulations, here we go!")
elif st.session_state['magic_unlocked'] is False:
    st.error("### Better luck next time...")
    st.image("https://images.unsplash.com/photo-1518020382113-a7e8fc38eac9?q=80&w=500", caption="Oh no! No magic today.", width=300)

# --- MAIN APP LOGIC ---
if st.session_state['magic_unlocked'] is True:
    
    BCA_BATCHES = ["BCA 2025", "BCA 2024", "BCA 2023"]
    MCA_BATCHES = ["MCA 2025", "MCA 2024"]
    # Per request: Focus specifically on the 75% threshold
    THRESHOLDS = [75]
    BLACKLIST = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SOFT SKILL", 
                 "SWIMMING", "ZUMBA", "FREESLOT", "TABLE TENNIS", "SS ATOM", "FREE SLOT"]

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

    def run_batch_logic(df, batches, writer, cols):
        batch_summaries = []
        for group in batches:
            course, year = group.split()
            mask = (df[cols['batch']].astype(str).str.contains(course, case=False, na=False)) & \
                   (df[cols['batch']].astype(str).str.contains(year, na=False))
            batch_df = df[mask].copy()
            batch_df = batch_df[~batch_df[cols['subject']].astype(str).str.contains('|'.join(BLACKLIST), case=False, na=False)]
            
            # Get unique sections within this batch
            sections = batch_df[cols['section']].unique() if cols['section'] in batch_df.columns else ["Default"]
            
            for section in sections:
                if section == "Default":
                    sec_df = batch_df
                else:
                    sec_df = batch_df[batch_df[cols['section']] == section]

                for limit in THRESHOLDS:
                    shortage_df = sec_df[sec_df[cols['attendance']] < limit].copy()
                    student_count = 0
                    if not shortage_df.empty:
                        grid = shortage_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['section']],
                                                        columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
                        grid.insert(0, 'Sl No.', range(1, len(grid) + 1))
                        
                        sub_cols = grid.columns[5:] # Adjusted for the inclusion of Section column
                        theory_cols = [c for c in sub_cols if "LAB" not in str(c).upper()]
                        grid['Theory Avg'] = grid[theory_cols].mean(axis=1).round(2)
                        grid['Final Avg'] = grid[sub_cols].mean(axis=1).round(2)
                        
                        sheet_name = f"{group} {section} <{limit}%"[:31]
                        grid.to_excel(writer, sheet_name=sheet_name, index=False)
                        apply_styles(writer.sheets[sheet_name])
                        
                        with st.expander(f"👁️ Section {section} - {group} (Below {limit}%)"):
                            subject_counts = grid[sub_cols].notna().sum().reset_index()
                            subject_counts.columns = ['Subject', 'Student Count']
                            fig_sub = px.bar(subject_counts, x='Subject', y='Student Count', 
                                             title=f"Shortage Distribution - {group} Section {section}",
                                             color='Student Count', color_continuous_scale='Reds')
                            st.plotly_chart(fig_sub, use_container_width=True)
                            st.dataframe(grid, hide_index=True, use_container_width=True)
                        
                        student_count = len(grid)
                    batch_summaries.append({'Batch': group, 'Section': section, 'Threshold': f"Below {limit}%", 'Count': student_count})
        return batch_summaries

    st.divider()
    uploaded_file = st.file_uploader("📂 Upload Attendance Excel to Start the Magic", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file, header=2)
        # Added Section column mapping (Assumed Column 7 based on your index pattern, update if different)
        cols = {'roll': df.columns[1], 'name': df.columns[2], 'batch': df.columns[6],
                'section': df.columns[7], 'subject': df.columns[8], 'attendance': df.columns[15]}
        
        df[cols['attendance']] = pd.to_numeric(df[cols['attendance']], errors='coerce')

        output = io.BytesIO()
        tab1, tab2, tab3 = st.tabs(["🎓 BCA Analytics", "📜 MCA Analytics", "📊 Executive Summary"])

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with tab1: bca_data = run_batch_logic(df, BCA_BATCHES, writer, cols)
            with tab2: mca_data = run_batch_logic(df, MCA_BATCHES, writer, cols)
            with tab3:
                summary_df = pd.DataFrame(bca_data + mca_data)
                # Grouping summary to show per batch and section
                st.subheader("📈 Section-wise Attendance Shortage (<75%)")
                
                fig = px.bar(summary_df, x='Batch', y='Count', color='Section', 
                             barmode='group', title="Shortage Count by Batch and Section",
                             labels={'Count': 'Students < 75% Attendance'},
                             color_discrete_sequence=px.colors.qualitative.Bold)
                st.plotly_chart(fig, use_container_width=True)
                
                st.divider()
                st.info("Master Summary Data Table")
                st.dataframe(summary_df, hide_index=True, use_container_width=True)
                
                summary_df.to_excel(writer, sheet_name='MASTER SUMMARY', index=False)
                apply_styles(writer.sheets['MASTER SUMMARY'], is_summary=True)

        st.divider()
        st.download_button(label="📥 Download Magic Report", data=output.getvalue(), file_name="Sectionwise_Attendance_Report.xlsx")

# Copyright Footer
st.markdown('<div class="footer">© VMS</div>', unsafe_allow_html=True)