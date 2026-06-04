import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob
import datetime

# Set page configuration to wide and beautiful theme
st.set_page_config(
    page_title="ระบบภาพรวมรายงานการผลิตรายเดือน",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Premium CSS to align exactly with the user's screenshots
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&family=Outfit:wght@300;400;500;600;700&display=swap');
        
        * {
            font-family: 'Sarabun', 'Outfit', sans-serif;
        }
        
        /* Main background color */
        .stApp {
            background-color: #f8fafc;
        }
        
        /* Sidebar layout styling */
        [data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #e2e8f0;
            padding: 10px;
        }
        
        /* Custom Sidebar Button styles */
        .sidebar-btn {
            display: block;
            width: 100%;
            text-align: left;
            background-color: white;
            color: #475569;
            border: 1px solid #e2e8f0;
            padding: 11px 15px;
            margin-bottom: 8px;
            border-radius: 10px;
            font-weight: 500;
            font-size: 0.92rem;
            text-decoration: none;
            transition: all 0.2s ease;
        }
        .sidebar-btn:hover {
            background-color: #fee2e2;
            color: #dc2626;
            border-color: #fca5a5;
            padding-left: 18px;
            text-decoration: none !important;
        }
        .sidebar-btn-active {
            display: block;
            width: 100%;
            text-align: left;
            background-color: #dc2626;
            color: white !important;
            border: 1px solid #dc2626;
            padding: 11px 15px;
            margin-bottom: 8px;
            border-radius: 10px;
            font-weight: 600;
            font-size: 0.92rem;
            text-decoration: none !important;
            box-shadow: 0 4px 8px rgba(220, 38, 38, 0.2);
        }
        
        /* Summary Cards in Overview page */
        .summary-card {
            background-color: white;
            border-radius: 16px;
            padding: 20px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02);
            transition: all 0.25s ease;
        }
        .summary-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 15px rgba(0, 0, 0, 0.05);
        }
        
        /* Force Streamlit horizontal blocks to not wrap on desktop viewports to maximize width usage */
        @media (min-width: 992px) {
            div[data-testid="stHorizontalBlock"] {
                flex-wrap: nowrap !important;
                gap: 20px !important;
            }
        }
    </style>
""", unsafe_allow_html=True)


PAGE_LABELS = {
    "overview": "📊 ภาพรวมการผลิต",
    "oven1": "🌽 เตาอบแป้ง 1",
    "oven2": "🌽 เตาอบแป้ง 2",
    "feed": "🌽 เตาอบกาก",
    "germ": "🌽 เตาอบเยอม",
    "meal": "🌽 เตาอบโปรตีน",
    "hsw": "💧 HSW"
}

# Determine the active page from session state or query parameters
if "active_page" not in st.session_state:
    st.session_state.active_page = "overview"

# Read query parameters for state persistence
params = st.query_params
if "page" in params:
    # Match query param back to menu keys
    param_val = params["page"]
    if param_val in PAGE_LABELS:
        st.session_state.active_page = param_val

# Sidebar Top Branding
st.sidebar.markdown("""
<div style="text-align: center; margin-bottom: 20px; padding-top: 10px;">
    <h2 style="margin: 0; color: #1e3a8a; font-size: 1.5rem; font-weight: 700;">📊 รายงานการผลิต</h2>
    <p style="margin: 0; font-size: 0.8rem; color: #64748b;">Monthly Production System</p>
</div>
""", unsafe_allow_html=True)

# Find Excel files in workspace directory
excel_files = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
if not excel_files:
    excel_files = [f for f in glob.glob("../*.xlsx") if not os.path.basename(f).startswith("~$")]

if excel_files:
    file_options = {os.path.basename(f): f for f in excel_files}
    selected_filename = st.sidebar.selectbox("📂 เลือกไฟล์ Excel", list(file_options.keys()))
    excel_path = file_options[selected_filename]
else:
    st.sidebar.warning("⚠️ ไม่พบไฟล์รายงาน Excel")
    excel_path = None

# Sidebar Menu (Wrapped in a beautiful solid Red Frame/Border exactly as requested)
st.sidebar.markdown('<div style="margin-top: 15px; margin-bottom: 10px;"></div>', unsafe_allow_html=True)

# Build buttons inside the red border container
buttons_html = ""
for key, label in PAGE_LABELS.items():
    active_class = "sidebar-btn-active" if st.session_state.active_page == key else "sidebar-btn"
    buttons_html += f'<a href="?page={key}" target="_self" class="{active_class}">{label}</a>'

st.sidebar.markdown(f"""
<div style="border: 2.5px solid #dc2626; border-radius: 14px; padding: 12px; background-color: #fef2f2; box-shadow: 0 4px 10px rgba(220, 38, 38, 0.1); margin-bottom: 20px;">
    <div style="text-align: center; color: #dc2626; font-weight: bold; font-size: 0.95rem; margin-bottom: 12px; border-bottom: 1.5px solid #fecaca; padding-bottom: 8px;">
        ⚙️ เตาอบและไลน์การผลิต
    </div>
    {buttons_html}
</div>
""", unsafe_allow_html=True)

# Excel Data Parser Function
def load_sheet_data(df, sheet_name):
    header_idx = None
    for idx, row in df.iterrows():
        if row.astype(str).str.contains("วันที่ผลิต").any():
            header_idx = idx
            break
            
    if header_idx is None:
        return None
        
    headers = df.iloc[header_idx].tolist()
    data_df = df.iloc[header_idx + 1:].copy()
    data_df.columns = headers
    
    # Drop columns that are completely null or have no name
    data_df = data_df.loc[:, data_df.columns.notna()]
    data_df.columns = [col.strip() for col in data_df.columns]
    
    # Drop rows where "วันที่ผลิต" is null
    data_df = data_df.dropna(subset=["วันที่ผลิต"])
    
    # Convert 'วันที่ผลิต' to datetime
    data_df["วันที่ผลิต"] = pd.to_datetime(data_df["วันที่ผลิต"], errors="coerce")
    data_df = data_df.dropna(subset=["วันที่ผลิต"])
    
    # Convert numerical columns
    for col in data_df.columns:
        if col != "วันที่ผลิต" and "หมายเหตุ" not in col:
            data_df[col] = pd.to_numeric(data_df[col], errors="coerce").fillna(0)
            
    # Calculate downtime in hours
    if "เวลาหยุดเครื่อง (นาที)" in data_df.columns:
        data_df["เวลาหยุดเครื่อง (ชั่วโมง)"] = data_df["เวลาหยุดเครื่อง (นาที)"] / 60.0
    else:
        data_df["เวลาหยุดเครื่อง (ชั่วโมง)"] = 0.0
        data_df["เวลาหยุดเครื่อง (นาที)"] = 0.0
        
    # Always recalculate yield deduction if columns exist
    deduct_cols = [col for col in data_df.columns if "ผลผลิตที่หัก" in col]
    mix_cols = [col for col in data_df.columns if "จำนวนผสม" in col]
    if deduct_cols and mix_cols:
        data_df[deduct_cols[0]] = data_df["ผลิตได้ (ตัน)"] - data_df[mix_cols[0]]
        
    # Always recalculate variance
    if "แผนผลิต (ตัน)" in data_df.columns:
        if deduct_cols:
            data_df["ส่วนต่าง (ตัน)"] = data_df[deduct_cols[0]] - data_df["แผนผลิต (ตัน)"]
        elif "ผลิตได้ (ตัน)" in data_df.columns:
            data_df["ส่วนต่าง (ตัน)"] = data_df["ผลิตได้ (ตัน)"] - data_df["แผนผลิต (ตัน)"]
        
    # Clean remarks column
    remarks_col = [col for col in data_df.columns if "หมายเหตุ" in col]
    if remarks_col:
        data_df["หมายเหตุ_สะอาด"] = data_df[remarks_col[0]].astype(str).str.strip().replace({"-": "", "nan": ""})
    else:
        data_df["หมายเหตุ_สะอาด"] = ""
        
    return data_df

def get_thai_name(sheet_name):
    SHEET_NAMES_TH = {
        'Oven1': 'เตาอบแป้ง 1',
        'Oven2': 'เตาอบแป้ง 2',
        'Feed': 'เตาอบกาก',
        'Germ': 'เตาอบเยอม',
        'Meal': 'เตาอบโปรตีน',
        'HSW': 'HSW'
    }
    for key, val in SHEET_NAMES_TH.items():
        if key.lower() in sheet_name.lower():
            return val
    return sheet_name

# Helper to apply live data editor edits to a active copy of df for dynamic UI updates
def apply_editor_changes(df, page_key):
    editor_key = f"editor_table_{page_key}"
    if editor_key in st.session_state:
        state = st.session_state[editor_key]
        if "edited_rows" in state and state["edited_rows"]:
            df = df.copy()
            for idx, changes in state["edited_rows"].items():
                for col, val in changes.items():
                    if col in df.columns:
                        df.iat[idx, df.columns.get_loc(col)] = val
            
            # Recalculate derived columns in df
            # 1. Yield deduction
            deduct_cols = [col for col in df.columns if "ผลผลิตที่หัก" in col]
            mix_cols = [col for col in df.columns if "จำนวนผสม" in col]
            if deduct_cols and mix_cols:
                df[deduct_cols[0]] = df["ผลิตได้ (ตัน)"] - df[mix_cols[0]]
            
            # 2. Variance
            if "แผนผลิต (ตัน)" in df.columns:
                if deduct_cols:
                    df["ส่วนต่าง (ตัน)"] = df[deduct_cols[0]] - df["แผนผลิต (ตัน)"]
                elif "ผลิตได้ (ตัน)" in df.columns:
                    df["ส่วนต่าง (ตัน)"] = df["ผลิตได้ (ตัน)"] - df["แผนผลิต (ตัน)"]
            
            # 3. Downtime in hours
            if "เวลาหยุดเครื่อง (นาที)" in df.columns:
                df["เวลาหยุดเครื่อง (ชั่วโมง)"] = df["เวลาหยุดเครื่อง (นาที)"] / 60.0
    return df

# Main Program Execution
if excel_path:
    try:
        xl = pd.ExcelFile(excel_path)
        all_sheets = xl.sheet_names
        
        # Load sheets data
        data_dict = {}
        for sheet in all_sheets:
            raw_df = xl.parse(sheet)
            cleaned_df = load_sheet_data(raw_df, sheet)
            if cleaned_df is not None:
                data_dict[sheet] = cleaned_df
                
        # Function to save updated data back to the specific Excel sheet using openpyxl
        def save_data_to_excel(excel_path, sheet_name, edited_df):
            import openpyxl
            import datetime
            wb = openpyxl.load_workbook(excel_path)
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"ไม่พบชีท {sheet_name} ในไฟล์ Excel")
            
            ws = wb[sheet_name]
            
            # Find headers row index (1-indexed)
            header_row_idx = None
            for r in range(1, 15):
                row_vals = [ws.cell(row=r, column=c).value for c in range(1, 15)]
                row_vals_str = [str(x) for x in row_vals if x is not None]
                if any("วันที่ผลิต" in x for x in row_vals_str):
                    header_row_idx = r
                    break
                    
            if header_row_idx is None:
                raise ValueError("ไม่พบหัวตาราง 'วันที่ผลิต' ในชีทนี้")
                
            # Map header name to column index
            header_cols = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=header_row_idx, column=c).value
                if val is not None:
                    header_cols[str(val).strip()] = c
                    
            date_col_idx = header_cols.get("วันที่ผลิต")
            if not date_col_idx:
                raise ValueError("ไม่พบคอลัมน์ 'วันที่ผลิต' ในไฟล์ Excel")
                
            # Map date string to row index
            excel_date_to_row = {}
            for r in range(header_row_idx + 1, ws.max_row + 1):
                date_val = ws.cell(row=r, column=date_col_idx).value
                if date_val is not None:
                    if isinstance(date_val, (datetime.datetime, datetime.date)):
                        date_parsed = date_val.strftime('%Y-%m-%d')
                    else:
                        try:
                            date_parsed = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                        except:
                            continue
                    excel_date_to_row[date_parsed] = r
                    
            # Write row by row
            for idx, row in edited_df.iterrows():
                row_date = row['วันที่ผลิต']
                if isinstance(row_date, str):
                    date_str = pd.to_datetime(row_date).strftime('%Y-%m-%d')
                else:
                    date_str = row_date.strftime('%Y-%m-%d')
                    
                sheet_row = excel_date_to_row.get(date_str)
                if sheet_row is not None:
                    # Write editable cells
                    cols_to_write = ['แผนผลิต (ตัน)', 'ผลิตได้ (ตัน)', 'ส่วนต่าง (ตัน)', 'เวลาหยุดเครื่อง (นาที)', 'หมายเหตุ']
                    for col_name in cols_to_write:
                        if col_name in header_cols and col_name in row:
                            col_idx = header_cols[col_name]
                            val = row[col_name]
                            if pd.isna(val) or val == 'nan':
                                val = None
                            if col_name == 'หมายเหตุ' and (val == '' or val is None):
                                val = '-'
                            ws.cell(row=sheet_row, column=col_idx, value=val)
                    
                    # Recalculate yield deduction if columns exist
                    deduct_target_col = None
                    mix_col = None
                    for col_name in header_cols:
                        if "ผลผลิตที่หัก" in col_name:
                            deduct_target_col = col_name
                        elif "จำนวนผสม" in col_name:
                            mix_col = col_name
                            
                    actual_val = float(row.get('ผลิตได้ (ตัน)', 0))
                    mix_val = 0.0
                    deduct_val = actual_val
                    
                    if deduct_target_col and mix_col:
                        mix_cell_val = ws.cell(row=sheet_row, column=header_cols[mix_col]).value
                        try:
                            mix_val = float(mix_cell_val) if mix_cell_val is not None else 0.0
                        except:
                            mix_val = 0.0
                        deduct_val = actual_val - mix_val
                        ws.cell(row=sheet_row, column=header_cols[deduct_target_col], value=deduct_val)
                        
                    # Recalculate variance or save user manual edit
                    if 'ส่วนต่าง (ตัน)' in header_cols:
                        if 'ส่วนต่าง (ตัน)' in row and not pd.isna(row['ส่วนต่าง (ตัน)']):
                            try:
                                diff_val = float(row['ส่วนต่าง (ตัน)'])
                            except:
                                plan_val = float(row.get('แผนผลิต (ตัน)', 0))
                                diff_val = (deduct_val - plan_val) if deduct_target_col else (actual_val - plan_val)
                        else:
                            plan_val = float(row.get('แผนผลิต (ตัน)', 0))
                            diff_val = (deduct_val - plan_val) if deduct_target_col else (actual_val - plan_val)
                        ws.cell(row=sheet_row, column=header_cols['ส่วนต่าง (ตัน)'], value=diff_val)

            wb.save(excel_path)
            wb.close()

        # Helper to retrieve sheet name and data based on page key
        def get_sheet_and_data_for_page(page_key):
            mapping = {
                "oven1": ["oven1", "เตาอบแป้ง 1", "เตาอบแป้ง1"],
                "oven2": ["oven2", "เตาอบแป้ง 2", "เตาอบแป้ง2"],
                "feed": ["feed", "เตาอบกาก", "กาก"],
                "meal": ["meal", "เตาอบโปรตีน", "โปรตีน", "protein"],
                "germ": ["germ", "เตาอบเยอม", "เยอม"],
                "hsw": ["hsw", "steep", "evaporator"]
            }
            if page_key not in mapping:
                return None, None
            kws = mapping[page_key]
            for sname, df in data_dict.items():
                for kw in kws:
                    if kw.lower() in sname.lower():
                        return sname, df
            return None, None

        # Helper to retrieve data based on page for backward compatibility
        def get_data_for_page(page_key):
            _, df = get_sheet_and_data_for_page(page_key)
            return df

        # Determine report month name for display
        month_str = "พฤษภาคม 2569" # Default fallback
        if "พฤษภาคม" in selected_filename:
            month_str = "พฤษภาคม 2569"
        elif "พฤษภาคม" in all_sheets[0] or "May" in all_sheets[0]:
            month_str = "พฤษภาคม 2569"
        # Extract from raw cell if available
        first_sheet = list(data_dict.keys())[0]
        # Look in original sheet name/cell
        
        # ==========================================
        # 1. PAGE: OVERVIEW / DASHBOARD
        # ==========================================
        if st.session_state.active_page == "overview":
            # Header matching Image 1 (without greeting)
            st.markdown(f"""
            <div style="background-color: white; border-radius: 16px; padding: 24px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.02); margin-bottom: 25px;">
                <h1 style="margin: 0; color: #1e3a8a; font-size: 2rem; font-weight: 700;">ภาพรวมระบบรายงานการผลิต</h1>
                <p style="margin: 5px 0 0 0; color: #64748b; font-size: 0.95rem;">ภาพรวมคลังผลผลิตการทำงานวันนี้ — 🟢 อัปเดตข้อมูลล่าสุดแล้ว</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Aggregate metrics
            total_planned = 0
            total_actual = 0
            total_downtime_mins = 0
            lines_count = len(data_dict)
            
            for s, df in data_dict.items():
                if 'แผนผลิต (ตัน)' in df.columns:
                    total_planned += df['แผนผลิต (ตัน)'].sum()
                if 'ผลิตได้ (ตัน)' in df.columns:
                    total_actual += df['ผลิตได้ (ตัน)'].sum()
                if 'เวลาหยุดเครื่อง (นาที)' in df.columns:
                    total_downtime_mins += df['เวลาหยุดเครื่อง (นาที)'].sum()
            
            achievement_rate = (total_actual / total_planned * 100) if total_planned > 0 else 0
            total_variance = total_actual - total_planned
            
            # 4 Summary Cards matching the layout of Image 1
            st.markdown(f"""
            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px;">
                <!-- Card 1: Total Lines -->
                <div class="summary-card">
                    <div style="color: #64748b; font-size: 0.85rem; font-weight: 500; display: flex; align-items: center; gap: 8px;">
                        <span style="background-color: #eff6ff; color: #2563eb; border-radius: 6px; padding: 4px 8px; font-weight: bold;">🏢</span> ไลน์การผลิตทั้งหมด
                    </div>
                    <div style="font-size: 1.8rem; font-weight: 700; color: #1e293b; margin-top: 10px;">{lines_count} ไลน์</div>
                    <div style="font-size: 0.75rem; color: #10b981; margin-top: 4px; font-weight: 600;">🔵 พร้อมใช้งานทุกเตาอบ</div>
                </div>
                <!-- Card 2: Planned vs Actual -->
                <div class="summary-card">
                    <div style="color: #64748b; font-size: 0.85rem; font-weight: 500; display: flex; align-items: center; gap: 8px;">
                        <span style="background-color: #ecfdf5; color: #10b981; border-radius: 6px; padding: 4px 8px; font-weight: bold;">📈</span> ผลผลิตสะสมทั้งหมด
                    </div>
                    <div style="font-size: 1.8rem; font-weight: 700; color: #1e293b; margin-top: 10px;">{total_actual:,.2f} ตัน</div>
                    <div style="font-size: 0.75rem; color: #64748b; margin-top: 4px; font-weight: 500;">เป้าหมายรวม: {total_planned:,.2f} ตัน</div>
                </div>
                <!-- Card 3: Achievement Rate -->
                <div class="summary-card">
                    <div style="color: #64748b; font-size: 0.85rem; font-weight: 500; display: flex; align-items: center; gap: 8px;">
                        <span style="background-color: #fffbeb; color: #f59e0b; border-radius: 6px; padding: 4px 8px; font-weight: bold;">🧩</span> อัตราความสำเร็จรวม
                    </div>
                    <div style="font-size: 1.8rem; font-weight: 700; color: #1e293b; margin-top: 10px;">{achievement_rate:.1f}%</div>
                    <div style="font-size: 0.75rem; color: {'#10b981' if total_variance >= 0 else '#ef4444'}; margin-top: 4px; font-weight: 600;">
                        ส่วนต่าง: {'+' if total_variance >= 0 else ''}{total_variance:,.2f} ตัน
                    </div>
                </div>
                <!-- Card 4: Total Downtime -->
                <div class="summary-card">
                    <div style="color: #64748b; font-size: 0.85rem; font-weight: 500; display: flex; align-items: center; gap: 8px;">
                        <span style="background-color: #fef2f2; color: #ef4444; border-radius: 6px; padding: 4px 8px; font-weight: bold;">⚠️</span> เวลาหยุดเครื่องรวม
                    </div>
                    <div style="font-size: 1.8rem; font-weight: 700; color: #dc2626; margin-top: 10px;">{total_downtime_mins/60.0:,.1f} ชม.</div>
                    <div style="font-size: 0.75rem; color: #ef4444; margin-top: 4px; font-weight: 600;">หยุดสะสม {total_downtime_mins:,.0f} นาที</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Row of Charts
            c1, c2 = st.columns([7, 3])
            
            with c1:
                # Merge daily production metrics for all lines
                daily_merge = []
                for s, df in data_dict.items():
                    df_slim = df[['วันที่ผลิต', 'แผนผลิต (ตัน)', 'ผลิตได้ (ตัน)']].copy()
                    daily_merge.append(df_slim)
                merged_daily = pd.concat(daily_merge).groupby('วันที่ผลิต').sum().reset_index()
                
                # Grouped Bar chart comparing Target vs Actual daily
                fig_daily = go.Figure()
                fig_daily.add_trace(go.Bar(
                    x=merged_daily['วันที่ผลิต'],
                    y=merged_daily['แผนผลิต (ตัน)'],
                    name='เป้าหมายรวม (Plan)',
                    marker_color='#1e3a8a',
                    marker_cornerradius=8
                ))
                fig_daily.add_trace(go.Bar(
                    x=merged_daily['วันที่ผลิต'],
                    y=merged_daily['ผลิตได้ (ตัน)'],
                    name='ผลิตได้จริงรวม (Actual)',
                    marker_color='#84cc16',
                    marker_cornerradius=8
                ))
                
                fig_daily.update_layout(
                    title='<b>แนวโน้มความเคลื่อนไหวการผลิตสะสมรายวัน (แผนงานเทียบผลิตจริง)</b>',
                    barmode='group',
                    bargap=0.3,
                    bargroupgap=0.1,
                    xaxis=dict(title='วันที่ผลิต', tickformat='%d %b'),
                    yaxis=dict(title='ผลผลิตรวม (ตัน)'),
                    legend_orientation="h",
                    legend=dict(x=0, y=-0.15),
                    margin=dict(l=20, r=20, t=50, b=20),
                    height=350,
                    plot_bgcolor='rgba(248, 250, 252, 0.8)',
                    paper_bgcolor='rgba(0,0,0,0)',
                )
                
                with st.container(border=True):
                    st.plotly_chart(fig_daily, use_container_width=True)
                
            with c2:
                # Donut Chart for proportion of total actual production by line
                shares = []
                for s, df in data_dict.items():
                    name = get_thai_name(s)
                    actual = df['ผลิตได้ (ตัน)'].sum() if 'ผลิตได้ (ตัน)' in df.columns else 0
                    shares.append({'ไลน์การผลิต': name, 'ผลผลิตสะสม (ตัน)': actual})
                shares_df = pd.DataFrame(shares)
                
                fig_donut = px.pie(
                    shares_df,
                    values='ผลผลิตสะสม (ตัน)',
                    names='ไลน์การผลิต',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Prism
                )
                fig_donut.update_layout(
                    title='<b>สัดส่วนผลผลิตตามไลน์ (ตัน)</b>',
                    margin=dict(l=10, r=10, t=40, b=10),
                    height=350,
                    legend=dict(orientation="h", x=0, y=-0.2),
                    paper_bgcolor='rgba(0,0,0,0)',
                )
                
                with st.container(border=True):
                    st.plotly_chart(fig_donut, use_container_width=True)
                
            # Bottom row of blocks
            c3, c4 = st.columns(2)
            
            with c3:
                # Ovens below target list (red warning list)
                with st.container(border=True):
                    st.markdown("<h3 style='margin-top:0; color: #dc2626; font-size: 1.1rem; border-bottom: 1.5px solid #fee2e2; padding-bottom: 8px;'>⚠️ ไลน์ผลิตที่ต่ำกว่าเป้าหมาย (Ovens Below Target)</h3>", unsafe_allow_html=True)
                    
                    below_target_html = ""
                    for s, df in data_dict.items():
                        name = get_thai_name(s)
                        plan = df['แผนผลิต (ตัน)'].sum() if 'แผนผลิต (ตัน)' in df.columns else 0
                        actual = df['ผลิตได้ (ตัน)'].sum() if 'ผลิตได้ (ตัน)' in df.columns else 0
                        diff = actual - plan
                        if diff < 0:
                            pct = (actual / plan * 100) if plan > 0 else 0
                            below_target_html += f"""
                            <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px 15px; background-color: #fff5f5; border: 1px solid #fee2e2; border-radius: 10px; margin-bottom: 10px;">
                                <div>
                                    <div style="font-weight: 600; color: #991b1b; font-size: 0.92rem;">{name}</div>
                                    <div style="font-size: 0.78rem; color: #7f1d1d;">ผลิตได้จริง {actual:,.2f} ตัน จากเป้า {plan:,.2f} ตัน</div>
                                </div>
                                <div style="text-align: right;">
                                    <div style="font-weight: 700; color: #dc2626; font-size: 0.95rem;">-{abs(diff):,.2f} ตัน</div>
                                    <div style="font-size: 0.78rem; color: #dc2626; font-weight: 600;">{pct:.1f}% ของเป้า</div>
                                </div>
                            </div>
                            """
                    if not below_target_html:
                        below_target_html = "<div style='color: #166534; font-weight: 600; text-align: center; padding: 50px 0;'>🎉 ทุกไลน์ผลิตได้ตรงตามเป้าหมาย (100% Target Hit)</div>"
                    
                    st.markdown(below_target_html, unsafe_allow_html=True)
                
            with c4:
                # Recent incidents timeline list
                with st.container(border=True):
                    st.markdown("<h3 style='margin-top:0; color: #1e3a8a; font-size: 1.1rem; border-bottom: 1.5px solid #e2e8f0; padding-bottom: 8px;'>🚨 เหตุการณ์หยุดเครื่องจักรล่าสุด (Recent Machine Incidents)</h3>", unsafe_allow_html=True)
                    
                    all_incidents = []
                    for s, df in data_dict.items():
                        name = get_thai_name(s)
                        inc = df[df['เวลาหยุดเครื่อง (นาที)'] > 0].copy()
                        for _, row in inc.iterrows():
                            all_incidents.append({
                                'line': name,
                                'date': row['วันที่ผลิต'],
                                'downtime': row['เวลาหยุดเครื่อง (นาที)'],
                                'remark': row['หมายเหตุ_สะอาด'] if row['หมายเหตุ_สะอาด'] else "ไม่มีบันทึกรายละเอียดปัญหา"
                            })
                    all_incidents = sorted(all_incidents, key=lambda x: x['date'], reverse=True)
                    
                    incidents_html = ""
                    for item in all_incidents[:4]:
                        date_str = item['date'].strftime('%d/%m/%Y')
                        incidents_html += f"""
                        <div style="display: flex; gap: 12px; padding: 10px 0; border-bottom: 1px solid #f1f5f9; align-items: flex-start;">
                            <div style="background-color: #fee2e2; color: #dc2626; border-radius: 8px; padding: 5px 8px; font-weight: 700; font-size: 0.82rem; text-align: center; min-width: 120px;">
                                {item['downtime']:.0f} นาที ({item['downtime']/60.0:.1f} ชม.)
                            </div>
                            <div>
                                <div style="font-weight: 600; color: #334155; font-size: 0.9rem;">{item['line']} — วันที่ {date_str}</div>
                                <div style="font-size: 0.8rem; color: #64748b; margin-top: 1px;">สาเหตุ: {item['remark']}</div>
                            </div>
                        </div>
                        """
                    if not incidents_html:
                        incidents_html = "<div style='color: #166534; font-weight: 600; text-align: center; padding: 50px 0;'>🎉 ไม่พบประวัติการหยุดทำงานของเครื่องจักร</div>"
                    
                    st.markdown(incidents_html, unsafe_allow_html=True)
                
        # ==========================================
        # 2. PAGE: OVEN-SPECIFIC VIEW (Oven1 - HSW)
        # ==========================================
        else:
            page_key = st.session_state.active_page
            page_title = PAGE_LABELS[page_key]
            
            # Fetch sheet name and data for selected oven
            sheet_name, df = get_sheet_and_data_for_page(page_key)
            
            if df is not None:
                # Apply live edits from the editor table to an active copy of df for reactive cards and charts
                df_active = apply_editor_changes(df, page_key)
                
                # 1. Header Card matching Image 3 (Linear gradient with pulse/waveform SVG)
                st.markdown(f"""
                <div style="display: flex; align-items: center; gap: 15px; background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); color: white; padding: 20px; border-radius: 16px; margin-bottom: 25px; box-shadow: 0 4px 10px rgba(30, 58, 138, 0.15);">
                    <div style="background-color: rgba(255, 255, 255, 0.2); border-radius: 12px; padding: 10px; display: flex; align-items: center; justify-content: center;">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" class="feather feather-activity" style="color: white;">
                            <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
                        </svg>
                    </div>
                    <div>
                        <h2 style="margin: 0; font-size: 1.8rem; font-weight: 700; color: white; line-height: 1.2;">Production - {page_title}</h2>
                        <p style="margin: 3px 0 0 0; opacity: 0.9; font-size: 0.95rem;">สรุปผลการผลิตประจำเดือน {month_str}</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Fetch statistics from active df
                target_sum = df_active['แผนผลิต (ตัน)'].sum() if 'แผนผลิต (ตัน)' in df_active.columns else 0
                actual_sum = df_active['ผลิตได้ (ตัน)'].sum() if 'ผลิตได้ (ตัน)' in df_active.columns else 0
                
                # Sum the manually edited variance if present, otherwise calculate it
                if 'ส่วนต่าง (ตัน)' in df_active.columns:
                    diff_sum = df_active['ส่วนต่าง (ตัน)'].sum()
                else:
                    diff_sum = actual_sum - target_sum
                    
                downtime_sum_mins = df_active['เวลาหยุดเครื่อง (นาที)'].sum() if 'เวลาหยุดเครื่อง (นาที)' in df_active.columns else 0
                
                diff_prefix = "+" if diff_sum >= 0 else ""
                diff_color = "#15803d" if diff_sum >= 0 else "#dc2626"
                
                # 2. Four KPI cards matching Image 3 structure
                st.markdown(f"""
                <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px;">
                    <!-- Card 1: Plan Target -->
                    <div style="background-color: white; border: 1.5px solid #e2e8f0; border-radius: 16px; padding: 18px 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.01);">
                        <div style="color: #1e3a8a; font-size: 0.95rem; font-weight: 600; margin-bottom: 12px;">เป้าหมายการผลิต (ตัน)</div>
                        <div style="display: flex; justify-content: center; align-items: center; gap: 12px; color: #1e3a8a; font-size: 2.2rem; font-weight: 700;">
                            <span style="font-size: 1.8rem;">📊</span> {target_sum:,.2f}
                        </div>
                    </div>
                    <!-- Card 2: Actual Output -->
                    <div style="background-color: white; border: 1.5px solid #e2e8f0; border-radius: 16px; padding: 18px 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.01);">
                        <div style="color: #15803d; font-size: 0.95rem; font-weight: 600; margin-bottom: 12px;">ยอดการผลิตได้ (ตัน)</div>
                        <div style="display: flex; justify-content: center; align-items: center; gap: 12px; color: #15803d; font-size: 2.2rem; font-weight: 700;">
                            <span style="font-size: 1.8rem;">📈</span> {actual_sum:,.2f}
                        </div>
                    </div>
                    <!-- Card 3: Variance -->
                    <div style="background-color: white; border: 1.5px solid #e2e8f0; border-radius: 16px; padding: 18px 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.01);">
                        <div style="color: #b45309; font-size: 0.95rem; font-weight: 600; margin-bottom: 12px;">ส่วนต่างการผลิต</div>
                        <div style="display: flex; justify-content: center; align-items: center; gap: 12px; color: {diff_color}; font-size: 2.2rem; font-weight: 700;">
                            <span style="font-size: 1.8rem;">🧩</span> {diff_prefix}{diff_sum:,.2f}
                        </div>
                    </div>
                    <!-- Card 4: Downtime minutes/hours -->
                    <div style="background-color: white; border: 1.5px solid #e2e8f0; border-radius: 16px; padding: 18px 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.01);">
                        <div style="color: #dc2626; font-size: 0.95rem; font-weight: 600; margin-bottom: 12px;">เวลาหยุดเครื่อง (นาที)</div>
                        <div style="display: flex; justify-content: center; align-items: center; gap: 12px; color: #dc2626; font-size: 2.2rem; font-weight: 700; line-height: 1;">
                            <span style="font-size: 1.8rem;">⚠️</span> {downtime_sum_mins:,.0f}
                        </div>
                        <div style="font-size: 0.8rem; color: #ef4444; margin-top: 4px; font-weight: 600;">({downtime_sum_mins/60.0:,.2f} ชั่วโมง)</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Column 1 (70%): Line trend, Column 2 (30%): Incident breakdown
                col_chart, col_reasons = st.columns([7, 3])
                
                with col_chart:
                    # Line chart matching colors and styles from Image 3
                    # Grouped Bar chart comparing Target vs Actual daily
                    fig_daily = go.Figure()
                    fig_daily.add_trace(go.Bar(
                        x=df_active['วันที่ผลิต'],
                        y=df_active['แผนผลิต (ตัน)'],
                        name='แผนผลิต (Plan)',
                        marker_color='#1e3a8a',
                        marker_cornerradius=8
                    ))
                    fig_daily.add_trace(go.Bar(
                        x=df_active['วันที่ผลิต'],
                        y=df_active['ผลิตได้ (ตัน)'],
                        name='ผลิตได้จริง (Actual)',
                        marker_color='#84cc16',
                        marker_cornerradius=8
                    ))
                    
                    fig_daily.update_layout(
                        title='<b>แนวโน้มการผลิตรายวัน (แผนงานเทียบผลิตจริง)</b>',
                        barmode='group',
                        bargap=0.3,
                        bargroupgap=0.1,
                        xaxis=dict(title='วันที่ผลิต', tickformat='%d %b'),
                        yaxis=dict(title='ผลผลิต (ตัน)'),
                        legend_orientation="h",
                        legend=dict(x=0, y=-0.15),
                        margin=dict(l=20, r=20, t=50, b=20),
                        height=350,
                        plot_bgcolor='rgba(248, 250, 252, 0.8)',
                        paper_bgcolor='rgba(0,0,0,0)',
                    )
                    
                    with st.container(border=True):
                        st.plotly_chart(fig_daily, use_container_width=True)
                    
                    # Calculate Monthly Key Stats for selected oven
                    best_day_idx = df_active['ผลิตได้ (ตัน)'].idxmax() if not df_active.empty and 'ผลิตได้ (ตัน)' in df_active.columns else None
                    if best_day_idx is not None and not pd.isna(best_day_idx):
                        best_day_row = df_active.loc[best_day_idx]
                        best_day_val = best_day_row['ผลิตได้ (ตัน)']
                        best_day_dt = best_day_row['วันที่ผลิต'].strftime('%d/%m/%Y')
                    else:
                        best_day_val = 0
                        best_day_dt = "-"
                        
                    days_met = (df_active['ผลิตได้ (ตัน)'] >= df_active['แผนผลิต (ตัน)']).sum() if 'แผนผลิต (ตัน)' in df_active.columns and 'ผลิตได้ (ตัน)' in df_active.columns else 0
                    total_days = len(df_active)
                    achievement_pct = (days_met / total_days * 100) if total_days > 0 else 0
                    
                    avg_actual = df_active['ผลิตได้ (ตัน)'].mean() if 'ผลิตได้ (ตัน)' in df_active.columns else 0
                    avg_plan = df_active['แผนผลิต (ตัน)'].mean() if 'แผนผลิต (ตัน)' in df_active.columns else 0
                    
                    avg_downtime = df_active['เวลาหยุดเครื่อง (นาที)'].mean() if 'เวลาหยุดเครื่อง (นาที)' in df_active.columns else 0
                    days_with_downtime = (df_active['เวลาหยุดเครื่อง (นาที)'] > 0).sum() if 'เวลาหยุดเครื่อง (นาที)' in df_active.columns else 0
                    
                    with st.container(border=True):
                        st.markdown("<h3 style='margin-top:0; color: #1e3a8a; font-size: 1.1rem; border-bottom: 1.5px solid #e2e8f0; padding-bottom: 8px;'>🏆 สถิติสำคัญประจำเดือน (Monthly Key Stats)</h3>", unsafe_allow_html=True)
                        st.markdown(f"""
                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 15px; margin-top: 10px;">
                            <!-- Best Day -->
                            <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 15px; display: flex; align-items: center; gap: 12px;">
                                <div style="background-color: #dcfce7; font-size: 1.8rem; border-radius: 10px; width: 45px; height: 45px; display: flex; align-items: center; justify-content: center; color: #15803d;">🌟</div>
                                <div>
                                    <div style="color: #166534; font-size: 0.8rem; font-weight: 600;">วันผลิตสูงสุด (Best Day)</div>
                                    <div style="font-size: 1.15rem; font-weight: 700; color: #14532d;">{best_day_val:,.2f} ตัน</div>
                                    <div style="font-size: 0.75rem; color: #166534;">วันที่ {best_day_dt}</div>
                                </div>
                            </div>
                            <!-- Target Achieved -->
                            <div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 12px; padding: 15px; display: flex; align-items: center; gap: 12px;">
                                <div style="background-color: #dbeafe; font-size: 1.8rem; border-radius: 10px; width: 45px; height: 45px; display: flex; align-items: center; justify-content: center; color: #1d4ed8;">🎯</div>
                                <div>
                                    <div style="color: #1e40af; font-size: 0.8rem; font-weight: 600;">บรรลุเป้าหมายผลิต (Target Achieved)</div>
                                    <div style="font-size: 1.15rem; font-weight: 700; color: #1e3a8a;">{days_met} จาก {total_days} วัน</div>
                                    <div style="font-size: 0.75rem; color: #1e40af;">คิดเป็น {achievement_pct:.1f}% ของวันทั้งหมด</div>
                                </div>
                            </div>
                            <!-- Avg Output -->
                            <div style="background-color: #fdf2f8; border: 1px solid #fbcfe8; border-radius: 12px; padding: 15px; display: flex; align-items: center; gap: 12px;">
                                <div style="background-color: #fce7f3; font-size: 1.8rem; border-radius: 10px; width: 45px; height: 45px; display: flex; align-items: center; justify-content: center; color: #be185d;">📈</div>
                                <div>
                                    <div style="color: #9d174d; font-size: 0.8rem; font-weight: 600;">ผลิตได้เฉลี่ยรายวัน (Avg Daily Output)</div>
                                    <div style="font-size: 1.15rem; font-weight: 700; color: #831843;">{avg_actual:,.2f} ตัน/วัน</div>
                                    <div style="font-size: 0.75rem; color: #9d174d;">(เป้าหมายเฉลี่ย: {avg_plan:,.2f} ตัน/วัน)</div>
                                </div>
                            </div>
                            <!-- Avg Downtime -->
                            <div style="background-color: #fff7ed; border: 1px solid #ffedd5; border-radius: 12px; padding: 15px; display: flex; align-items: center; gap: 12px;">
                                <div style="background-color: #ffedd5; font-size: 1.8rem; border-radius: 10px; width: 45px; height: 45px; display: flex; align-items: center; justify-content: center; color: #c2410c;">⏱️</div>
                                <div>
                                    <div style="color: #9a3412; font-size: 0.8rem; font-weight: 600;">เวลาหยุดเครื่องเฉลี่ย (Avg Downtime)</div>
                                    <div style="font-size: 1.15rem; font-weight: 700; color: #7c2d12;">{avg_downtime:,.1f} นาที</div>
                                    <div style="font-size: 0.75rem; color: #9a3412;">เกิดเหตุ {days_with_downtime} วันในเดือนนี้</div>
                                </div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                
                with col_reasons:
                    # Donut Chart for proportion of downtime reasons for selected page
                    reasons_df = df_active[df_active['เวลาหยุดเครื่อง (นาที)'] > 0].copy()
                    if not reasons_df.empty:
                        # Clean/normalize remarks
                        reasons_df['หมายเหตุ_สะอาด'] = reasons_df['หมายเหตุ_สะอาด'].replace({"": "ไม่ระบุสาเหตุ", "nan": "ไม่ระบุสาเหตุ", None: "ไม่ระบุสาเหตุ"})
                        grouped_reasons = reasons_df.groupby('หมายเหตุ_สะอาด')['เวลาหยุดเครื่อง (นาที)'].sum().reset_index()
                        
                        fig_dt_donut = px.pie(
                            grouped_reasons,
                            values='เวลาหยุดเครื่อง (นาที)',
                            names='หมายเหตุ_สะอาด',
                            hole=0.4,
                            color_discrete_sequence=px.colors.qualitative.Reds
                        )
                        fig_dt_donut.update_layout(
                            title='<b>สัดส่วนสาเหตุหยุดเครื่อง (นาที)</b>',
                            margin=dict(l=10, r=10, t=40, b=10),
                            height=280,
                            legend=dict(orientation="h", x=0, y=-0.2),
                            paper_bgcolor='rgba(0,0,0,0)',
                        )
                        with st.container(border=True):
                            st.plotly_chart(fig_dt_donut, use_container_width=True)
                            
                    with st.container(border=True):
                        st.markdown("<h3 style='margin-top:0; color: #dc2626; font-size: 1.1rem; border-bottom: 1.5px solid #fee2e2; padding-bottom: 8px;'>🚨 ประวัติเครื่องจักรขัดข้อง</h3>", unsafe_allow_html=True)
                        if not reasons_df.empty:
                            # Show list of incidents sorted by downtime minutes descending
                            reasons_sorted = reasons_df.sort_values(by='เวลาหยุดเครื่อง (นาที)', ascending=False)
                            incidents_html = ""
                            for _, row in reasons_sorted.iterrows():
                                date_str = row['วันที่ผลิต'].strftime('%d/%m/%Y')
                                remark = row['หมายเหตุ_สะอาด'] if row['หมายเหตุ_สะอาด'] else "ไม่ระบุสาเหตุ"
                                incidents_html += f"""
                                <div style="display: flex; gap: 12px; padding: 10px 0; border-bottom: 1px solid #f1f5f9; align-items: flex-start;">
                                    <div style="background-color: #fee2e2; color: #dc2626; border-radius: 8px; padding: 4px 6px; font-weight: 700; font-size: 0.8rem; text-align: center; min-width: 90px;">
                                        {row['เวลาหยุดเครื่อง (นาที)']:.0f} นาที
                                        <br><span style="font-size: 0.68rem; font-weight: 500;">({row['เวลาหยุดเครื่อง (นาที)']/60.0:.1f} ชม.)</span>
                                    </div>
                                    <div>
                                        <div style="font-weight: 600; color: #334155; font-size: 0.85rem;">วันที่ {date_str}</div>
                                        <div style="font-size: 0.78rem; color: #64748b; margin-top: 1px;">สาเหตุ: {remark}</div>
                                    </div>
                                </div>
                                """
                            st.markdown(incidents_html, unsafe_allow_html=True)
                        else:
                            st.markdown("<div style='color: #166534; font-weight: 600; text-align: center; padding: 40px 0; font-size: 0.9rem;'>🎉 ไม่พบรายงานประวัติเครื่องหยุดขัดข้อง</div>", unsafe_allow_html=True)
                
                # 3. Middle: Interactive Daily Detail Viewer
                st.markdown("### 🔍 เจาะลึกข้อมูลรายวัน (Daily Detail Viewer)")
                
                # Dropdown for selecting day
                day_options = df_active.sort_values('วันที่ผลิต')['วันที่ผลิต'].dt.strftime('%d (%Y-%m-%d)').tolist()
                selected_day_str = st.selectbox("เลือกวันที่ผลิตที่ต้องการดูรายละเอียด:", day_options)
                
                selected_date_parsed = pd.to_datetime(selected_day_str.split('(')[1].replace(')', '').strip())
                day_row = df_active[df_active['วันที่ผลิต'] == selected_date_parsed].iloc[0]
                
                plan_day = day_row.get('แผนผลิต (ตัน)', 0)
                actual_day = day_row.get('ผลิตได้ (ตัน)', 0)
                diff_day = day_row.get('ส่วนต่าง (ตัน)', actual_day - plan_day)
                dt_day = day_row.get('เวลาหยุดเครื่อง (นาที)', 0)
                remark_day = day_row.get('หมายเหตุ_สะอาด', "")
                if not remark_day or remark_day == "nan":
                    remark_day = "ไม่มีเหตุเครื่องขัดข้องหรือบันทึกหมายเหตุ"
                    
                diff_day_color = "#10b981" if diff_day >= 0 else "#ef4444"
                diff_day_prefix = "+" if diff_day >= 0 else ""
                
                with st.container(border=True):
                    st.markdown(f"""
                    <h4 style="margin-top: 0; color: #1e3a8a; font-size: 1.1rem; border-bottom: 1px solid #f1f5f9; padding-bottom: 8px;">
                        📝 เจาะลึกและแก้ไขข้อมูลประจำวันผลิตที่: {selected_date_parsed.strftime('%d/%m/%Y')}
                    </h4>
                    """, unsafe_allow_html=True)
                    
                    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                    with col_e1:
                        new_plan = st.number_input(
                            "เป้าหมายแผนงาน (ตัน)",
                            value=float(plan_day),
                            step=0.1,
                            key=f"plan_day_in_{page_key}_{selected_date_parsed.strftime('%Y%m%d')}"
                        )
                    with col_e2:
                        new_actual = st.number_input(
                            "ผลผลิตทำได้จริง (ตัน)",
                            value=float(actual_day),
                            step=0.1,
                            key=f"actual_day_in_{page_key}_{selected_date_parsed.strftime('%Y%m%d')}"
                        )
                    with col_e3:
                        new_diff = new_actual - new_plan
                        diff_day_color = "#10b981" if new_diff >= 0 else "#ef4444"
                        diff_day_prefix = "+" if new_diff >= 0 else ""
                        st.markdown(f"""
                        <div style="padding-top: 5px;">
                            <div style="color: #64748b; font-size: 0.8rem; font-weight: 500;">ส่วนต่างวันผลิต (คำนวณ)</div>
                            <div style="font-size: 1.3rem; font-weight: 700; color: {diff_day_color}; margin-top: 5px;">
                                {diff_day_prefix}{new_diff:.2f} ตัน
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col_e4:
                        new_dt = st.number_input(
                            "เวลาหยุดเครื่องจักร (นาที)",
                            value=float(dt_day),
                            step=1.0,
                            key=f"dt_day_in_{page_key}_{selected_date_parsed.strftime('%Y%m%d')}"
                        )
                        
                    new_remark = st.text_input(
                        "บันทึกข้อมูล/ปัญหาประจำวัน (หมายเหตุ):",
                        value="" if remark_day == "ไม่มีเหตุเครื่องขัดข้องหรือบันทึกหมายเหตุ" or remark_day == "-" else remark_day,
                        key=f"remark_day_in_{page_key}_{selected_date_parsed.strftime('%Y%m%d')}"
                    )
                    
                    if st.button("💾 บันทึกข้อมูลเฉพาะวันนี้ (Save Daily Entry)", key=f"btn_save_day_{page_key}_{selected_date_parsed.strftime('%Y%m%d')}", type="primary"):
                        single_row_df = pd.DataFrame([{
                            'วันที่ผลิต': selected_date_parsed,
                            'แผนผลิต (ตัน)': new_plan,
                            'ผลิตได้ (ตัน)': new_actual,
                            'เวลาหยุดเครื่อง (นาที)': new_dt,
                            'หมายเหตุ': new_remark
                        }])
                        try:
                            save_data_to_excel(excel_path, sheet_name, single_row_df)
                            st.success(f"💾 บันทึกข้อมูลวันที่ {selected_date_parsed.strftime('%d/%m/%Y')} ลงไฟล์ Excel สำเร็จแล้ว! กำลังโหลดข้อมูลใหม่...")
                            st.rerun()
                        except Exception as ex:
                            st.error(f"เกิดข้อผิดพลาดในการบันทึกข้อมูล: {str(ex)}")
 
                
                # 4. Bottom: Detailed daily table ("ขอรายละเอียด รายวันขึ้นข้อมูลด้วย และจัดให้ดูสวยงามครับ")
                st.markdown("### 📋 รายละเอียดข้อมูลรายวัน (Daily Log)")
                
                # Format dataframe with colors
                display_df = df.copy()
                display_df['วันที่ผลิต'] = display_df['วันที่ผลิต'].dt.strftime('%Y-%m-%d')
                
                # Format numeric columns for display
                for col in display_df.columns:
                    if col not in ['วันที่ผลิต', 'หมายเหตุ', 'หมายเหตุ_สะอาด']:
                        display_df[col] = display_df[col].apply(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)
                
                # We style values for presentation
                # Format columns list to display nicely
                cols_order = ['วันที่ผลิต', 'แผนผลิต (ตัน)', 'ผลิตได้ (ตัน)', 'ส่วนต่าง (ตัน)', 'เวลาหยุดเครื่อง (นาที)', 'เวลาหยุดเครื่อง (ชั่วโมง)', 'หมายเหตุ']
                cols_to_render = [c for c in cols_order if c in display_df.columns]
                
                with st.container(border=True):
                    # Set up data editor with disabled calculated fields (ส่วนต่าง (ตัน) is now EDITABLE)
                    disabled_cols = ['วันที่ผลิต', 'เวลาหยุดเครื่อง (ชั่วโมง)']
                    
                    # Prepare dataframe copy to prevent representation float issues
                    df_to_edit = df[cols_to_render].copy()
                    if 'วันที่ผลิต' in df_to_edit.columns:
                        df_to_edit['วันที่ผลิต'] = df_to_edit['วันที่ผลิต'].dt.date
                    for c in df_to_edit.select_dtypes(include=['float64', 'float32']).columns:
                        df_to_edit[c] = df_to_edit[c].round(2)
                        
                    edited_df = st.data_editor(
                        df_to_edit,
                        disabled=[c for c in disabled_cols if c in cols_to_render],
                        use_container_width=True,
                        hide_index=True,
                        key=f"editor_table_{page_key}"
                    )
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("💾 บันทึกการแก้ไขในตารางทั้งหมด (Save All Table Changes)", key=f"btn_save_table_{page_key}"):
                            try:
                                save_data_to_excel(excel_path, sheet_name, edited_df)
                                st.success("💾 บันทึกการแก้ไขข้อมูลตารางทั้งหมดลงไฟล์ Excel สำเร็จเรียบร้อยแล้ว! กำลังโหลดข้อมูลใหม่...")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"เกิดข้อผิดพลาดในการบันทึกข้อมูลตาราง: {str(ex)}")
                    with col_b2:
                        try:
                            with open(excel_path, "rb") as f:
                                file_data = f.read()
                            st.download_button(
                                label="📥 ดาวน์โหลดไฟล์ Excel อัปเดตล่าสุด (Download Excel)",
                                data=file_data,
                                file_name=selected_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"btn_download_table_{page_key}"
                            )
                        except Exception as ex:
                            st.error(f"ไม่สามารถเตรียมไฟล์ดาวน์โหลดได้: {str(ex)}")

            else:
                st.error("❌ ไม่พบข้อมูลสำหรับไลน์การผลิตนี้")
                
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {str(e)}")
        st.exception(e)
else:
    st.info("กรุณานำไฟล์รายงานการผลิต (สกุลไฟล์ .xlsx) มาใส่ไว้ในโฟลเดอร์นี้เพื่อแสดงผลภาพรวม")
