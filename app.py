import streamlit as st
import openpyxl
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
from datetime import datetime

# 1. ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(
    page_title="ระบบสร้างใบวางบิลอัจฉริยะ", 
    page_icon="🌽", 
    layout="wide"
)

# 2. ตกแต่งด้วย Custom CSS
st.markdown("""
    <style>
    .main { background-color: #0f172a; }
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        padding: 24px;
        border-radius: 16px;
        border: 1px solid #475569;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
    }
    .header-title { color: #f8fafc; font-size: 28px; font-weight: 700; margin-bottom: 8px; }
    .header-subtitle { color: #94a3b8; font-size: 15px; }
    .stButton>button { border-radius: 10px; font-weight: 600; transition: all 0.2s ease; }
    div[data-baseweb="checkbox"] {
        padding: 8px 12px;
        background-color: #1e293b;
        border-radius: 8px;
        margin-bottom: 6px;
        border: 1px solid #334155;
    }
    </style>
""", unsafe_allow_html=True)

def add_highlight(run):
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement('w:highlight')
    highlight.set(qn('w:val'), 'yellow')
    rPr.append(highlight)

def format_date(val):
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    elif val:
        return str(val).split()[0]
    return "-"

# --- ส่วนหัวของเว็บ ---
st.markdown("""
    <div class="header-card">
        <div class="header-title">🌽 ระบบสร้างใบวางบิลข้าวโพด (รองรับทุกโครงสร้าง Excel)</div>
        <div class="header-subtitle">แปลงข้อมูลตาราง Excel เป็นเอกสาร Word อัตโนมัติ ยืดหยุ่น รองรับไฟล์ทุกรูปแบบ</div>
    </div>
""", unsafe_allow_html=True)

# --- แถบเครื่องมือด้านข้าง (Sidebar) ---
with st.sidebar:
    st.header("⚙️ เมนูตั้งค่า")
    uploaded_file = st.file_uploader("1. อัปโหลดไฟล์ Excel", type=["xlsx"])
    st.markdown("---")
    st.caption("ระบบอ่านหัวตารางอัตโนมัติ")

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    raw_sheet_names = wb.sheetnames
    sheet_options = {s.strip(): s for s in raw_sheet_names}
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_display = st.selectbox("📌 เลือกชีทที่ต้องการ:", list(sheet_options.keys()))
    
    if selected_display:
        actual_sheet_name = sheet_options[selected_display]
        ws = wb[actual_sheet_name]
        
        # 1. ค้นหาแถวที่เป็นหัวตาราง (Header Row)
        header_row = -1
        col_map = {}
        
        for r in range(1, min(20, ws.max_row + 1)):
            row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
            # ตรวจหาคอลัมน์สำคัญ
            for c_idx, val in enumerate(row_vals, 1):
                if "ทะเบียน" in val:
                    header_row = r
                if val:
                    col_map[val] = c_idx
            if header_row != -1:
                break
                
        if header_row == -1:
            header_row = 4 # Default แถว 4
            
        # สร้างรายชื่อคอลัมน์ทั้งหมดที่พบในแถวหัวตาราง
        headers_found = {}
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=header_row, column=c).value or f"คอลัมน์ {c}").strip()
            headers_found[val] = c
            
        header_keys = list(headers_found.keys())
        
        # ฟังก์ชั่นช่วยเดาหาคอลัมน์อัตโนมัติ
        def guess_col(keywords, default_col_idx):
            for k, idx in headers_found.items():
                if any(kw in k for kw in keywords):
                    return k
            return header_keys[default_col_idx - 1] if len(header_keys) >= default_col_idx else header_keys[0]

        # 2. เมนูตั้งค่าจับคู่คอลัมน์ (เผื่อผู้ใช้ต้องการเปลี่ยนเอง)
        with st.expander("🛠️ ตรวจสอบการจับคู่คอลัมน์ (คลิกเพื่อแก้ไขหากข้อมูลไม่ตรง)", expanded=False):
            st.info("ระบบเดาหัวตารางให้อัตโนมัติ หากข้อมูลขึ้นไม่ตรง สามารถปรับเปลี่ยนคอลัมน์ได้จากตัวเลือกด้านล่างครับ")
            c1, c2, c3, c4 = st.columns(4)
            sel_date_up = c1.selectbox("คอลัมน์ วันที่ขึ้นสินค้า:", header_keys, index=header_keys.index(guess_col(["ขึ้นสินค้า", "วันที่ขึ้น"], 2)))
            sel_plate = c2.selectbox("คอลัมน์ ทะเบียนรถ:", header_keys, index=header_keys.index(guess_col(["ทะเบียน"], 4)))
            sel_dest = c3.selectbox("คอลัมน์ ปลายทาง:", header_keys, index=header_keys.index(guess_col(["ปลายทาง"], 8)))
            sel_price = c4.selectbox("คอลัมน์ ราคา:", header_keys, index=header_keys.index(guess_col(["ราคา"], 13)))
            
            c5, c6, c7, _ = st.columns(4)
            sel_w_start = c5.selectbox("คอลัมน์ นน.ต้นทาง:", header_keys, index=header_keys.index(guess_col(["ต้นทาง", "นน.ต้นทาง"], 10)))
            sel_w_end = c6.selectbox("คอลัมน์ นน.ปลายทาง:", header_keys, index=header_keys.index(guess_col(["ปลายทาง", "นน.ปลายทาง"], 11)))
            sel_w_diff = c7.selectbox("คอลัมน์ ส่วนต่างน้ำหนัก:", header_keys, index=header_keys.index(guess_col(["ส่วนต่าง"], 12)))

        idx_date_up = headers_found[sel_date_up]
        idx_plate = headers_found[sel_plate]
        idx_dest = headers_found[sel_dest]
        idx_price = headers_found[sel_price]
        idx_w_start = headers_found[sel_w_start]
        idx_w_end = headers_found[sel_w_end]
        idx_w_diff = headers_found[sel_w_diff]

        # 3. อ่านข้อมูลจากตาราง
        items_data = []
        for row in range(header_row + 1, ws.max_row + 1):
            plate = ws.cell(row=row, column=idx_plate).value
            
            if plate and str(plate).strip() not in ["ทะเบียน", "None", "", "รวม"]:
                date_up = ws.cell(row=row, column=idx_date_up).value
                destination = ws.cell(row=row, column=idx_dest).value
                weight_start = ws.cell(row=row, column=idx_w_start).value
                weight_end = ws.cell(row=row, column=idx_w_end).value
                weight_diff = ws.cell(row=row, column=idx_w_diff).value
                price = ws.cell(row=row, column=idx_price).value
                
                # ถ้าน้ำหนักปลายทางไม่มี ให้ใช้น้ำหนักต้นทางแทน
                display_weight = weight_end if (weight_end is not None and str(weight_end).strip() != "") else weight_start
                
                items_data.append({
                    "date_up": format_date(date_up),
                    "plate": str(plate).strip(),
                    "destination": str(destination).strip() if destination else "-",
                    "weight_start": weight_start,
                    "weight_end": weight_end,
                    "display_weight": display_weight,
                    "weight_diff": weight_diff,
                    "price": price
                })

        if len(items_data) == 0:
            st.warning(f"⚠️ ไม่พบข้อมูลรายการในชีท **{selected_display}**")
        else:
            all_dates = sorted(list(set([item["date_up"] for item in items_data if item["date_up"] != "-"])))
            date_filter_options = ["แสดงทั้งหมด"] + all_dates
            
            with col_sel2:
                selected_date = st.selectbox("📅 กรองตามวันที่ขึ้นสินค้า:", date_filter_options)
            
            if selected_date != "แสดงทั้งหมด":
                filtered_items = [item for item in items_data if item["date_up"] == selected_date]
            else:
                filtered_items = items_data

            m1, m2 = st.columns(2)
            m1.metric(label="ชีทที่เลือก", value=selected_display)
            m2.metric(label="จำนวนรายการทั้งหมด", value=f"{len(filtered_items)} รายการ")
            
            st.write("---")
            st.subheader("📋 เลือกรายการที่ต้องการส่งออก")
            
            def set_all_checkboxes(status):
                for idx in range(1, len(filtered_items) + 1):
                    st.session_state[f"chk_{selected_date}_{idx}"] = status

            btn_c1, btn_c2, _ = st.columns([1, 1, 2])
            btn_c1.button("✅ เลือกทั้งหมด", on_click=set_all_checkboxes, args=(True,), use_container_width=True)
            btn_c2.button("❌ ไม่เลือกเลย", on_click=set_all_checkboxes, args=(False,), use_container_width=True)
            
            st.write("")
            selected_indices = []
            
            for idx, item in enumerate(filtered_items, 1):
                key_name = f"chk_{selected_date}_{idx}"
                if key_name not in st.session_state:
                    st.session_state[key_name] = True
                    
                w_val = item['display_weight']
                w_str = f"{w_val:,.0f}" if isinstance(w_val, (int, float)) else f"{w_val or '-'}"
                
                label = f"ข้อ {idx} [{item['date_up']}]: ทะเบียน {item['plate']} | ปลายทาง: {item['destination']} | นน.: {w_str} | ราคา: {item['price'] or '-'}"
                
                is_selected = st.checkbox(label, key=key_name)
                if is_selected:
                    selected_indices.append(item)
            
            st.write("---")
            
            if st.button("🚀 สร้างไฟล์ Word (.docx)", type="primary", use_container_width=True):
                if len(selected_indices) == 0:
                    st.error("กรุณาเลือกอย่างน้อย 1 รายการก่อนสร้างเอกสาร")
                else:
                    doc = Document()
                    for section in doc.sections:
                        section.page_width = Inches(8.27)
                        section.page_height = Inches(11.69)
                        section.top_margin = Inches(1.0)
                        section.bottom_margin = Inches(1.0)
                        section.left_margin = Inches(1.2)
                        section.right_margin = Inches(1.2)
                        
                    p_title = doc.add_paragraph()
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_title.paragraph_format.space_after = Pt(24)
                    r_title = p_title.add_run(f"วางบิลข้าวโพด เจ้นัชชา ({selected_display})")
                    r_title.font.name = "TH Sarabun PSK"
                    r_title.font.size = Pt(22)
                    r_title.font.bold = True
                    
                    for count, item in enumerate(selected_indices, 1):
                        if count > 1:
                            doc.add_page_break()
                            
                        header_text = f"ทะเบียน {item['plate']} ({selected_display}-{item['destination']})"
                        
                        weight_val = item['display_weight']
                        if item['weight_end'] is not None and str(item['weight_end']).strip() != "":
                            weight_str = f"นน.ปลายทาง {weight_val:,.0f}" if isinstance(weight_val, (int, float)) else f"นน.ปลายทาง {weight_val or '-'}"
                        else:
                            weight_str = f"นน.ต้นทาง {weight_val:,.0f}" if isinstance(weight_val, (int, float)) else f"นน.ต้นทาง {weight_val or '-'}"
                        
                        diff_val = item['weight_diff']
                        if isinstance(diff_val, (int, float)):
                            sign = "+" if diff_val > 0 else ""
                            diff_str = f"ส่วนต่างน้ำหนัก {sign}{diff_val:,.0f} กิโลกรัม"
                        elif diff_val is not None and str(diff_val).strip() != "":
                            diff_str = f"ส่วนต่างน้ำหนัก {diff_val}"
                        else:
                            diff_str = "ส่วนต่างน้ำหนัก -"
                            
                        price_str = f"ราคา {item['price']}" if item['price'] else "ราคา -"
                        
                        for text in [header_text, weight_str, diff_str, price_str]:
                            p = doc.add_paragraph()
                            p.paragraph_format.space_after = Pt(12)
                            r = p.add_run(text)
                            r.font.name = "TH Sarabun PSK"
                            r.font.size = Pt(16)
                            
                        p_status = doc.add_paragraph()
                        p_status.paragraph_format.space_after = Pt(12)
                        r_status = p_status.add_run("*ลงเรียบร้อย*")
                        r_status.font.name = "TH Sarabun PSK"
                        r_status.font.size = Pt(16)
                        r_status.font.bold = True
                        add_highlight(r_status)
                    
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)
                    
                    st.success(f"สร้างไฟล์สำเร็จ! สรุปเลือกมาทำทั้งหมด {len(selected_indices)} รายการ")
                    st.download_button(
                        label="📥 กดดาวน์โหลดไฟล์ Word (.docx)",
                        data=bio,
                        file_name=f"วางบิลข้าวโพด_{selected_display}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ Excel ทางเมนูด้านซ้ายเพื่อเริ่มต้นใช้งาน")