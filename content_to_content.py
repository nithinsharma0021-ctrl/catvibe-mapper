import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import zipfile
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 1. SMART HEADER TRANSLATOR
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

# --- UI START ---
st.set_page_config(page_title="CatVibe | Bulk Mapper", layout="wide")
st.title("⚡ CatVibe Bulk Smart Mapper")
st.markdown("Upload your Myntra Templates and Seller Update files below to cross-reference attributes automatically.")

# File Uploaders
template_files = st.file_uploader("1. Upload Target Myntra Templates (.xlsx) [Multiple Allowed]", type=["xlsx"], accept_multiple_files=True)
seller_files = st.file_uploader("2. Upload Seller Files (.xlsx, .csv) [Multiple Allowed]", type=["xlsx", "csv"], accept_multiple_files=True)

if st.button("🚀 Run Bulk Mapping") and template_files and seller_files:
    with st.spinner("Processing files and mapping attributes..."):
        master_style_dict = {}
        
        # 2. PARSE ALL SELLER FILES INTO MASTER DICTIONARY
        for uploaded_seller in seller_files:
            try:
                # A. Handle CSV Files
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

                # B. Handle Excel Files (Auto-Detect Multi-Row vs Simple Table)
                xls = pd.ExcelFile(uploaded_seller)
                sheet_name = 'Style Sheet' if 'Style Sheet' in xls.sheet_names else ('Styles' if 'Styles' in xls.sheet_names else xls.sheet_names[0])
                df_raw = pd.read_excel(uploaded_seller, sheet_name=sheet_name, header=None)
                
                # Check for multi-row complex header layout
                is_complex_header = False
                for r in range(min(4, len(df_raw))):
                    row_str = " ".join([str(x).lower() for x in df_raw.iloc[r].dropna().values])
                    if 'field names' in row_str or 'current inputs' in row_str or 'correct inputs' in row_str or 'new inputs' in row_str:
                        is_complex_header = True
                        break

                if is_complex_header:
                    # Multi-row seller header logic: STRICTLY PRIORITIZE NEW DATA
                    h_row = 0 if 'field names' in " ".join([str(x).lower() for x in df_raw.iloc[0].dropna().values]) else 2
                    sh_row = h_row + 1
                    start_row = sh_row + 1
                    
                    headers = df_raw.iloc[h_row].ffill().tolist()
                    sub_headers = df_raw.iloc[sh_row].fillna('').tolist()
                    
                    combined_cols = []
                    for i, (h, sh) in enumerate(zip(headers, sub_headers)):
                        if i == 0:
                            combined_cols.append('style_id')
                            continue
                        h_str = str(h).strip().lower().replace('\n', ' ') if pd.notna(h) else ''
                        sh_str = str(sh).strip().lower().replace('\n', ' ')
                        # Mark columns based on current vs new distinction
                        if 'correct' in sh_str or 'new' in sh_str: combined_cols.append(f"{h_str}_new")
                        elif 'current' in sh_str: combined_cols.append(f"{h_str}_current")
                        else: combined_cols.append(f"{h_str}_{sh_str}")
                            
                    df_data = df_raw.iloc[start_row:].copy()
                    df_data.columns = combined_cols
                    df_data['style_id'] = df_data['style_id'].ffill()
                    
                    attr_key_col = next((c for c in df_data.columns if c.startswith('attribute') and not c.endswith('new') and not c.endswith('current')), None)
                    attr_val_col = next((c for c in df_data.columns if c.startswith('attribute') and c.endswith('new')), None)
                    spec_col = next((c for c in df_data.columns if 'specification_new' in c), None)

                    for _, row in df_data.iterrows():
                        sid = clean_id(row['style_id'])
                        if not sid or sid == 'nan': continue
                        if sid not in master_style_dict: master_style_dict[sid] = {}

                        # Multiline Cells (New Data Focus)
                        if attr_key_col and spec_col and pd.notna(row[attr_key_col]) and pd.notna(row[spec_col]) and '\n' in str(row[attr_key_col]):
                            attrs = str(row[attr_key_col]).split('\n')
                            vals = str(row[spec_col]).split('\n')
                            for i in range(min(len(attrs), len(vals))):
                                key = map_attribute_header(attrs[i])
                                val = vals[i].strip()
                                if key and val and val.lower() != 'nan': master_style_dict[sid][key] = val

                        # Multi-Row (New Data Focus)
                        elif attr_key_col and attr_val_col and pd.notna(row[attr_key_col]) and pd.notna(row[attr_val_col]):
                            k = str(row[attr_key_col]).strip()
                            v = str(row[attr_val_col]).strip()
                            if k.lower() != 'nan' and v.lower() != 'nan': master_style_dict[sid][map_attribute_header(k)] = v

                        # Standard Columns (ONLY process _new columns, ignore _current)
                        for col in df_data.columns:
                            if col.endswith('_new') and col != spec_col and not col.startswith('attribute'):
                                if pd.notna(row[col]) and str(row[col]).strip().lower() != 'nan':
                                    key = map_attribute_header(col.replace('_new', ''))
                                    master_style_dict[sid][key] = str(row[col]).strip()

                else:
                    # Simple Table (Direct Mapping of Singular True Values)
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

        st.success(f"✅ Loaded updates for {len(master_style_dict)} unique styles from seller files.")

        # 3. INJECT INTO TEMPLATES
        zip_buffer = io.BytesIO()
        total_mapped = 0
        total_branded = 0
        
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
                
                total_mapped += mapped_count
                total_branded += brand_injection_count
                
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                zip_file.writestr(f"Mapped_{uploaded_template.name}", wb_buffer.getvalue())

        # 4. DOWNLOAD RESULTS
        st.info(f"🎉 Processing Complete! Updated {total_mapped} total attributes across {len(template_files)} template(s).")
        if total_branded > 0:
            st.warning(f"🛡️ Auto-injected missing Brand Name into {total_branded} titles.")

        if len(template_files) == 1:
            wb_buffer.seek(0)
            st.download_button(
                label="💾 Download Updated Myntra Template",
                data=wb_buffer.getvalue(),
                file_name=f"Mapped_{template_files[0].name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.download_button(
                label="📦 Download All Updated Templates (.zip)",
                data=zip_buffer.getvalue(),
                file_name="CatVibe_Bulk_Templates.zip",
                mime="application/zip"
            )
