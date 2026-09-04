import streamlit as st
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
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

# ฟังก์ชั่นใส่เส้นขอบตาราง Word
def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

# ฟังก์ชั่นจัดรูปแบบวันที่ + แก้ไขปี 1969 อัตโนมัติ
def format_date(val):
    if not val or str(val).strip() in ["-", "None", ""]:
        return "-"
        
    if isinstance(val, datetime):
        year = val.year
        if year == 1969:
            year = 2026
        elif year > 2500:
            year = year - 543
        return val.strftime(f"%d/%m/{year}")
    
    val_str = str(val).strip().split()[0]
    if "1969" in val_str:
        val_str = val_str.replace("1969", "2026")
    elif "/69" in val_str:
        val_str = val_str.replace("/69", "/2026")
        
    return val_str

def parse_date_for_sort(date_str):
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except Exception:
        return datetime.min

# --- ส่วนหัวของเว็บ ---
st.markdown("""
    <div class="header-card">
        <div class="header-title">🌽 ระบบสร้างใบวางบิลข้าวโพด (รูปแบบตาราง เจพีที)</div>
        <div class="header-subtitle">แปลงข้อมูลตาราง Excel ออกเป็นใบวางบิลตาราง Word แบบมาตรฐาน</div>
    </div>
""", unsafe_allow_html=True)

# --- แถบเครื่องมือด้านข้าง (Sidebar) ---
with st.sidebar:
    st.header("⚙️ เมนูตั้งค่า")
    uploaded_file = st.file_uploader("1. อัปโหลดไฟล์ Excel", type=["xlsx"])
    st.markdown("---")
    st.caption("ระบบดึงข้อมูลลงตารางให้อัตโนมัติ")

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
        
        # ค้นหาแถวที่เป็นหัวตาราง
        header_row = 4
        for r in range(1, min(15, ws.max_row + 1)):
            row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
            if any(kw in val for val in row_vals for kw in ["ทะเบียน", "ลำดับ", "สินค้า"]):
                header_row = r
                break
                
        headers_found = {}
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=header_row, column=c).value or "").strip()
            if val:
                headers_found[val] = c
            else:
                headers_found[f"คอลัมน์ {c}"] = c
            
        header_keys = list(headers_found.keys())
        
        def find_exact_col(exact_keywords, fallback_idx):
            for kw in exact_keywords:
                for k in header_keys:
                    if k == kw:
                        return k
            for kw in exact_keywords:
                for k in header_keys:
                    if kw in k and "ส่วนต่าง" not in k and "จำนวน" not in k:
                        return k
            if len(header_keys) >= fallback_idx:
                return header_keys[fallback_idx - 1]
            return header_keys[0]

        default_date_up = find_exact_col(["วัน/เดือน/ปี ขึ้นสินค้า", "วันที่ขึ้นสินค้า", "วัน/เดือน/ปี", "วันที่"], 2)
        default_date_down = find_exact_col(["วัน/เดือน/ปี ลงสินค้า", "วันที่ลงสินค้า", "วันที่ลง"], 3)
        default_plate = find_exact_col(["ทะเบียน"], 4)
        default_prod = find_exact_col(["สินค้า"], 5)
        default_origin = find_exact_col(["ต้นทาง"], 7)
        default_dest = find_exact_col(["ปลายทาง"], 8)
        default_price = find_exact_col(["ราคา"], 13)
        default_w_start = find_exact_col(["นน.ต้นทาง", "น้ำหนักต้นทาง"], 10)
        default_w_end = find_exact_col(["นน.ปลายทาง", "น้ำหนักปลายทาง"], 11)

        idx_date_up = headers_found[default_date_up]
        idx_date_down = headers_found[default_date_down]
        idx_plate = headers_found[default_plate]
        idx_prod = headers_found.get(default_prod, idx_plate)
        idx_origin = headers_found.get(default_origin, idx_plate)
        idx_dest = headers_found[default_dest]
        idx_price = headers_found[default_price]
        idx_w_start = headers_found[default_w_start]
        idx_w_end = headers_found[default_w_end]

        # อ่านข้อมูลจากตาราง
        items_data = []
        last_valid_date = "-"
        
        for row in range(header_row + 1, ws.max_row + 1):
            plate = ws.cell(row=row, column=idx_plate).value
            
            if plate and str(plate).strip() not in ["ทะเบียน", "None", "", "รวม"]:
                raw_date_up = ws.cell(row=row, column=idx_date_up).value
                raw_date_down = ws.cell(row=row, column=idx_date_down).value
                
                formatted_d_up = format_date(raw_date_up)
                formatted_d_down = format_date(raw_date_down)
                
                if formatted_d_up and formatted_d_up != "-":
                    last_valid_date = formatted_d_up
                
                product = ws.cell(row=row, column=idx_prod).value or "ข้าวโพด"
                origin = ws.cell(row=row, column=idx_origin).value or selected_display
                destination = ws.cell(row=row, column=idx_dest).value or "-"
                
                weight_start = ws.cell(row=row, column=idx_w_start).value or "-"
                weight_end = ws.cell(row=row, column=idx_w_end).value or "-"
                price = ws.cell(row=row, column=idx_price).value or "-"
                
                items_data.append({
                    "date_up": last_valid_date,
                    "date_down": formatted_d_down,
                    "plate": str(plate).strip(),
                    "product": str(product).strip(),
                    "origin": str(origin).strip(),
                    "destination": str(destination).strip(),
                    "weight_start": weight_start,
                    "weight_end": weight_end,
                    "price": price
                })

        if len(items_data) == 0:
            st.warning(f"⚠️ ไม่พบข้อมูลรายการในชีท **{selected_display}**")
        else:
            unique_dates = list(set([item["date_up"] for item in items_data if item["date_up"] != "-"]))
            sorted_dates = sorted(unique_dates, key=parse_date_for_sort)
            date_filter_options = ["แสดงทั้งหมด"] + sorted_dates
            
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
                    
                w_val = item['weight_end'] if item['weight_end'] != "-" else item['weight_start']
                w_str = f"{w_val:,.0f}" if isinstance(w_val, (int, float)) else f"{w_val}"
                
                label = f"ข้อ {idx} [{item['date_up']}]: ทะเบียน {item['plate']} | ปลายทาง: {item['destination']} | นน.: {w_str} | ราคา: {item['price']}"
                
                is_selected = st.checkbox(label, key=key_name)
                if is_selected:
                    selected_indices.append(item)
            
            st.write("---")
            
            # ปุ่มสร้างเอกสารแบบตารางใบวางบิล
            if st.button("🚀 สร้างไฟล์ Word ใบวางบิล (.docx)", type="primary", use_container_width=True):
                if len(selected_indices) == 0:
                    st.error("กรุณาเลือกอย่างน้อย 1 รายการก่อนสร้างเอกสาร")
                else:
                    doc = Document()
                    
                    # ตั้งค่าขอบกระดาษ 0.6 นิ้วตามแบบต้นฉบับ
                    for section in doc.sections:
                        section.page_width = Inches(8.27)
                        section.page_height = Inches(11.69)
                        section.top_margin = Inches(0.6)
                        section.bottom_margin = Inches(0.6)
                        section.left_margin = Inches(0.6)
                        section.right_margin = Inches(0.6)
                        
                    # 1. หัวเอกสาร
                    p_title = doc.add_paragraph()
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_title.paragraph_format.space_after = Pt(4)
                    r_title = p_title.add_run("ใบวางบิล")
                    r_title.font.name = "TH SarabunPSK"
                    r_title.font.size = Pt(24)
                    r_title.font.bold = True
                    
                    today_str = datetime.now().strftime("%d/%m/%Y")
                    p_date = doc.add_paragraph()
                    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_date.paragraph_format.space_after = Pt(16)
                    r_date = p_date.add_run(f"วันที่ {today_str}")
                    r_date.font.name = "TH SarabunPSK"
                    r_date.font.size = Pt(16)
                    
                    # 2. สร้างตารางรายการหลัก (11 คอลัมน์)
                    # เพิ่มแถวเผื่อให้มีบรรทัดว่าง 4 บรรทัดเหมือนต้นฉบับ
                    total_table_rows = max(len(selected_indices) + 2, 5)
                    table = doc.add_table(rows=total_table_rows, cols=11)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    set_table_borders(table)
                    
                    headers = ["ลำดับที่", "ทะเบียน", "สินค้า", "วันที่ขึ้น", "สถานที่ขึ้น", "สถานที่ลง", "วันที่ลง", "นน.\nต้นทาง", "นน.\nปลายทาง", "บาท/\nตัน", "ราคาสุทธิ"]
                    
                    # หัวตาราง
                    hdr_cells = table.rows[0].cells
                    for i, head_text in enumerate(headers):
                        p = hdr_cells[i].paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run(head_text)
                        r.font.name = "TH SarabunPSK"
                        r.font.size = Pt(11)
                        r.font.bold = True
                        
                    # ใส่ข้อมูลรายการ
                    total_price_sum = 0
                    for idx, item in enumerate(selected_indices, 1):
                        row_cells = table.rows[idx].cells
                        
                        w_start_str = f"{item['weight_start']:,.0f}" if isinstance(item['weight_start'], (int, float)) else str(item['weight_start'])
                        w_end_str = f"{item['weight_end']:,.0f}" if isinstance(item['weight_end'], (int, float)) else str(item['weight_end'])
                        
                        row_data = [
                            str(idx),
                            item['plate'],
                            item['product'],
                            item['date_up'],
                            item['origin'],
                            item['destination'],
                            item['date_down'],
                            w_start_str,
                            w_end_str,
                            str(item['price']),
                            "-"
                        ]
                        
                        for c_idx, val in enumerate(row_data):
                            p = row_cells[c_idx].paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 1, 3, 6] else WD_ALIGN_PARAGRAPH.LEFT
                            r = p.add_run(val)
                            r.font.name = "TH SarabunPSK"
                            r.font.size = Pt(11)
                            
                    # แถวสุดท้าย: ยอดสุทธิ (รวมเซลล์ 0-9)
                    sum_row_cells = table.rows[-1].cells
                    sum_row_cells[0].merge(sum_row_cells[9])
                    
                    p_sum_label = sum_row_cells[0].paragraphs[0]
                    p_sum_label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r_sum_label = p_sum_label.add_run("ยอดสุทธิ   ")
                    r_sum_label.font.name = "TH SarabunPSK"
                    r_sum_label.font.size = Pt(12)
                    r_sum_label.font.bold = True
                    
                    p_sum_val = sum_row_cells[10].paragraphs[0]
                    p_sum_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_sum_val = p_sum_val.add_run("-")
                    r_sum_val.font.name = "TH SarabunPSK"
                    r_sum_val.font.size = Pt(12)
                    r_sum_val.font.bold = True
                    
                    # 3. รายละเอียดการชำระเงินด้านล่าง
                    doc.add_paragraph().paragraph_format.space_after = Pt(8)
                    
                    p_pay_title = doc.add_paragraph()
                    p_pay_title.paragraph_format.space_after = Pt(4)
                    r_pay_title = p_pay_title.add_run("รายละเอียดการชำระเงิน")
                    r_pay_title.font.name = "TH SarabunPSK"
                    r_pay_title.font.size = Pt(14)
                    r_pay_title.font.bold = True
                    
                    pay_table = doc.add_table(rows=3, cols=2)
                    pay_data = [
                        ("ชื่อบัญชี :", "หจก. ทีเอ็นพี เลขที่ 44 หมู่ที่ 9 ต.ม่วงคำ"),
                        ("ธนาคาร :", "กสิกรไทย"),
                        ("เลขที่บัญชี :", "097-1-01627-2")
                    ]
                    
                    for r_idx, (label, val) in enumerate(pay_data):
                        row_cells = pay_table.rows[r_idx].cells
                        
                        p_lbl = row_cells[0].paragraphs[0]
                        r_lbl = p_lbl.add_run(label)
                        r_lbl.font.name = "TH SarabunPSK"
                        r_lbl.font.size = Pt(13)
                        r_lbl.font.bold = True
                        
                        p_val = row_cells[1].paragraphs[0]
                        r_val = p_val.add_run(val)
                        r_val.font.name = "TH SarabunPSK"
                        r_val.font.size = Pt(13)
                    
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)
                    
                    st.success(f"สร้างใบวางบิลสำเร็จ! รวม {len(selected_indices)} รายการ")
                    st.download_button(
                        label="📥 กดดาวน์โหลดไฟล์ใบวางบิล (.docx)",
                        data=bio,
                        file_name=f"ใบวางบิล_{selected_display}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ Excel ทางเมนูด้านซ้ายเพื่อเริ่มต้นใช้งาน")