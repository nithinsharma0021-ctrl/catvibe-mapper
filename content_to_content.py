import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import zipfile
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 1. PAGE CONFIG & INDIAN SIGN-ART MINIMALIST STYLING
st.set_page_config(page_title="C 2 C | Content Mapper", layout="wide", page_icon="⚡")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&family=Rozha+One&display=swap');

    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }

    .main-header {
        font-family: 'Rozha One', serif;
        font-size: 3.2rem;
        color: #111827;
        margin-bottom: 0.1rem;
        letter-spacing: 1px;
    }

    .sub-header {
        font-family: 'Poppins', sans-serif;
        font-size: 0.95rem;
        color: #B45309;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-bottom: 1.5rem;
    }

    /* Minimalist Card Containers */
    div[data-testid="stFileUploader"] {
        background-color: #FDFBF7;
        border: 2px solid #E5E7EB;
        border-radius: 8px;
        padding: 12px;
    }

    /* Primary Action Button - Bold Marigold Accent */
    div.stButton > button:first-child {
        background-color: #D97706;
        color: #FFFFFF;
        font-family: 'Poppins', sans-serif;
        font-weight: 700;
        font-size: 1rem;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        border: none;
        padding: 0.65rem 2rem;
        border-radius: 6px;
        transition: all 0.2s ease-in-out;
    }

    div.stButton > button:first-child:hover {
        background-color: #B45309;
        color: #FFFFFF;
        box-shadow: 0 4px 12px rgba(180, 83, 9, 0.25);
    }
</style>
""", unsafe_allow_html=True)

# 2. SMART HEADER TRANSLATOR
def map_attribute_header(raw_attr):
    raw = str(raw_attr).lower().strip()
    if 'display' in raw or 'title' in raw or raw == 'style name': return 'productdisplayname'
    if 'list view' in raw: return 'listviewname'
    if 'product detail' in raw: return 'productdetails'
    if 'size' in raw and 'fit' in raw: return 'sizeandfitdescription'
    if 'material' in raw and 'care' in raw: return 'materialcaredescription'
    if 'colour' in raw or 'color' in raw: return 'colour'
    if raw == 'patterns': return 'pattern'
    if raw == 'fashion trends': return 'maintrend'
    return re.sub(r'[^a-z0-9]', '', raw)

def clean_id(val):
    try: return str(int(float(str(val).strip())))
    except: return str(val).strip()

def normalize_col(name):
    name = str(name).strip().lower()
    return re.sub(r'[^a-z0-9]', '', name)

# --- HEADER SECTION ---
st.markdown('<div class="main-header">C 2 C</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Content to Content Smart Bulk Engine</div>', unsafe_allow_html=True)

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)

with col1:
    template_files = st.file_uploader(
        "1. Target Myntra Templates (.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )

with col2:
    seller_files = st.file_uploader(
        "2. Source Seller Files (.xlsx, .csv)", 
        type=["xlsx", "csv"], 
        accept_multiple_files=True
    )

st.markdown("---")

# --- EXECUTION ENGINE ---
if st.button("🚀 Run C 2 C Mapping") and template_files and seller_files:
    with st.spinner("Analyzing seller inputs & executing precision mapping..."):
        master_style_dict = {}
        
        # 3. PARSE SELLER FILES
        for uploaded_seller in seller_files:
            try:
                # CSV Format
                if uploaded_seller.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_seller)
                    style_col = next((c for c in df.columns if normalize_col(c) in ['styleid', 'style', 'id']), df.columns[0])
                    for _, row in df.iterrows():
                        sid = clean_id(row[style_col])
                        if sid and sid != 'nan':
                            if sid not in master_style_dict: master_style_dict[sid] = {}
                            for col in df.columns:
                                if col != style_col and pd.notna(row[col]):
                                    master_style_dict[sid][map_attribute_header(col)] = str(row[col]).strip()
                    continue

                # Excel Format
                xls = pd.ExcelFile(uploaded_seller)
                sheet_name = 'Style Sheet' if 'Style Sheet' in xls.sheet_names else ('Styles' if 'Styles' in xls.sheet_names else xls.sheet_names[0])
                df_raw = pd.read_excel(uploaded_seller, sheet_name=sheet_name, header=None)
                
                is_complex_header = False
                for r in range(min(4, len(df_raw))):
                    row_str = " ".join([str(x).lower() for x in df_raw.iloc[r].dropna().values])
                    if 'field names' in row_str or 'current inputs' in row_str or 'correct inputs' in row_str or 'new inputs' in row_str:
                        is_complex_header = True
                        break

                if is_complex_header:
                    start_row = 4 if len(df_raw) > 4 else 2
                    for r in range(start_row, len(df_raw)):
                        row = df_raw.iloc[r]
                        raw_sid = row[0]
                        if pd.isna(raw_sid): continue
                        sid = clean_id(raw_sid)
                        if not sid or sid == 'nan': continue
                        
                        if sid not in master_style_dict: master_style_dict[sid] = {}
                        
                        # Attribute column fallback (Col 10 = Name, Col 11 = Current, Col 12 = Correct/New)
                        if 10 < len(row) and pd.notna(row[10]):
                            attr_name = str(row[10]).strip()
                            val_curr_spec = row[11] if 11 < len(row) else None
                            val_new_spec = row[12] if 12 < len(row) else None
                            
                            target_val = None
                            if pd.notna(val_new_spec) and str(val_new_spec).strip().lower() not in ['nan', '']:
                                target_val = str(val_new_spec).strip()
                            elif pd.notna(val_curr_spec) and str(val_curr_spec).strip().lower() not in ['nan', '']:
                                target_val = str(val_curr_spec).strip()
                                
                            if attr_name and target_val:
                                master_style_dict[sid][map_attribute_header(attr_name)] = target_val

                        # Paired Columns (Display Name, List View, Product Details, Fit, Care)
                        paired_cols = [(4, 5, 4), (6, 7, 6), (8, 9, 8), (13, 14, 13), (15, 16, 15)]
                        for pair_curr, pair_new, header_idx in paired_cols:
                            if header_idx < df_raw.shape[1]:
                                h_val = df_raw.iloc[2, header_idx] if pd.notna(df_raw.iloc[2, header_idx]) else df_raw.iloc[1, header_idx]
                                if pd.notna(h_val):
                                    attr_header = str(h_val).strip()
                                    v_new = row[pair_new] if pair_new < len(row) else None
                                    v_curr = row[pair_curr] if pair_curr < len(row) else None
                                    
                                    final_v = None
                                    if pd.notna(v_new) and str(v_new).strip().lower() not in ['nan', '']:
                                        final_v = str(v_new).strip()
                                    elif pd.notna(v_curr) and str(v_curr).strip().lower() not in ['nan', '']:
                                        final_v = str(v_curr).strip()
                                        
                                    if attr_header and final_v:
                                        master_style_dict[sid][map_attribute_header(attr_header)] = final_v

                else:
                    # Simple 1-row Header Table
                    df_simple = pd.read_excel(uploaded_seller, sheet_name=sheet_name)
                    style_col = next((c for c in df_simple.columns if normalize_col(c) in ['styleid', 'style', 'id']), df_simple.columns[0])
                    
                    for _, row in df_simple.iterrows():
                        sid = clean_id(row[style_col])
                        if not sid or sid == 'nan': continue
                        if sid not in master_style_dict: master_style_dict[sid] = {}
                        
                        for col in df_simple.columns:
                            if col != style_col and pd.notna(row[col]):
                                val = str(row[col]).strip()
                                if val and val.lower() != 'nan':
                                    master_style_dict[sid][map_attribute_header(col)] = val

            except Exception as e:
                st.error(f"Error parsing {uploaded_seller.name}: {e}")

        st.success(f"✅ Extracted updates for **{len(master_style_dict)}** unique style IDs from seller files.")

        # 4. INJECT INTO TEMPLATES & AUDIT LOG
        zip_buffer = io.BytesIO()
        total_mapped = 0
        total_branded = 0
        audit_records = []
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for uploaded_template in template_files:
                wb = openpyxl.load_workbook(uploaded_template)
                mapped_count = 0
                brand_injection_count = 0

                for sheet_name in wb.sheetnames:
                    if sheet_name == 'masterdata': continue
                    ws = wb[sheet_name]
                    
                    col_map = {normalize_col(cell.value): idx for idx, cell in enumerate(ws[1], 1) if cell.value}
                    target_style_col = col_map.get('styleid')
                    brand_col_idx = col_map.get('brand')
                    
                    if target_style_col:
                        for row in range(2, ws.max_row + 1):
                            raw_sid = ws.cell(row=row, column=target_style_col).value
                            if not raw_sid: continue
                            sid = clean_id(raw_sid)
                            
                            if sid in master_style_dict:
                                source_data = master_style_dict[sid]
                                for mapped_attr, val in source_data.items():
                                    if mapped_attr in col_map:
                                        final_val = str(val).strip()
                                        
                                        # Auto-Brand Injector
                                        if mapped_attr == 'productdisplayname' and brand_col_idx:
                                            target_brand = str(ws.cell(row=row, column=brand_col_idx).value).strip()
                                            if target_brand and target_brand.lower() not in ['none', 'nan', '']:
                                                if target_brand.lower() not in final_val.lower():
                                                    final_val = f"{target_brand} {final_val}"
                                                    brand_injection_count += 1
                                        
                                        ws.cell(row=row, column=col_map[mapped_attr]).value = final_val
                                        mapped_count += 1
                                        
                                        # Audit trail
                                        orig_col_header = ws.cell(row=1, column=col_map[mapped_attr]).value
                                        audit_records.append({
                                            "Template File": uploaded_template.name,
                                            "Sheet Name": sheet_name,
                                            "Style ID": sid,
                                            "Attribute Mapped": orig_col_header,
                                            "Updated Value": final_val
                                        })
                
                total_mapped += mapped_count
                total_branded += brand_injection_count
                
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                zip_file.writestr(f"C2C_Mapped_{uploaded_template.name}", wb_buffer.getvalue())

        # 5. RESULTS & DOWNLOAD
        st.info(f"🎯 **C 2 C Mapping Complete!** Successfully written **{total_mapped}** attribute cells across template(s).")
        if total_branded > 0:
            st.warning(f"🛡️ Auto-injected Brand Name into **{total_branded}** titles.")

        # Interactive Audit Log
        if audit_records:
            df_audit = pd.DataFrame(audit_records)
            with st.expander("📊 Detailed Mapping Audit Log (Line-by-Line)", expanded=True):
                st.dataframe(df_audit, use_container_width=True)

        # Download Actions
        if len(template_files) == 1:
            wb_buffer.seek(0)
            st.download_button(
                label="💾 Download Updated C 2 C Template",
                data=wb_buffer.getvalue(),
                file_name=f"C2C_Mapped_{template_files[0].name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.download_button(
                label="📦 Download All C 2 C Templates (.zip)",
                data=zip_buffer.getvalue(),
                file_name="C2C_Mapped_Templates.zip",
                mime="application/zip"
            )
