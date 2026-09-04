import streamlit as st
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io
from datetime import datetime, date

# 1. ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(
    page_title="ระบบสร้างใบวางบิลข้าวโพด", 
    page_icon="🌽", 
    layout="wide"
)

# 2. Custom CSS
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

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def format_date_thai(val):
    if not val or str(val).strip() in ["-", "None", ""]:
        return "-"
        
    if isinstance(val, (datetime, date)):
        year = val.year
        if year == 1969 or year < 2000:
            year = 2026
        year_be = year + 543 if year < 2500 else year
        return f"{val.day}/{val.month}/{year_be}"
    
    val_str = str(val).strip().split()[0]
    if "1969" in val_str:
        val_str = val_str.replace("1969", "2569")
    elif "/69" in val_str:
        val_str = val_str.replace("/69", "/2569")
        
    return val_str

def add_word_field_formula(cell, formula_str, default_value_str, is_bold=False, font_size=14):
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), f'{formula_str} \\# "#,##0.00"')
    
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'TH SarabunPSK')
    rFonts.set(qn('w:hAnsi'), 'TH SarabunPSK')
    rPr.append(rFonts)
    
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(font_size * 2))
    rPr.append(sz)
    
    if is_bold:
        rPr.append(OxmlElement('w:b'))
        
    r.append(rPr)
    
    t = OxmlElement('w:t')
    t.text = default_value_str
    r.append(t)
    
    fldSimple.append(r)
    p._p.append(fldSimple)

# --- ส่วนหัวของเว็บ ---
st.markdown("""
    <div class="header-card">
        <div class="header-title">🌽 ระบบสร้างใบวางบิลข้าวโพด</div>
        <div class="header-subtitle">แปลงข้อมูล Excel เป็น Word แนวนอน - ตารางชำระเงินขนาดพอดีสายตา</div>
    </div>
""", unsafe_allow_html=True)

# --- แถบเครื่องมือด้านข้าง (Sidebar) ---
with st.sidebar:
    st.header("⚙️ เมนูตั้งค่า")
    uploaded_file = st.file_uploader("1. อัปโหลดไฟล์ Excel", type=["xlsx"])
    st.markdown("---")
    bill_date_input = st.date_input("2. เลือกวันที่ใบวางบิล:", value=date.today())

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    raw_sheet_names = wb.sheetnames
    sheet_options = {s.strip(): s for s in raw_sheet_names}
    
    selected_display = st.selectbox("📌 เลือกชีทที่ต้องการ:", list(sheet_options.keys()))
    
    if selected_display:
        actual_sheet_name = sheet_options[selected_display]
        ws = wb[actual_sheet_name]
        
        header_row = 1
        for r in range(1, min(15, ws.max_row + 1)):
            row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
            if any("ทะเบียน" in val or "สินค้า" in val or "ต้นทาง" in val for val in row_vals):
                header_row = r
                break
                
        headers_found = {}
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=header_row, column=c).value or "").strip()
            if val:
                headers_found[f"คอลัมน์ {c}: {val}"] = c
            else:
                headers_found[f"คอลัมน์ {c}: [ว่าง]"] = c
            
        header_options = list(headers_found.keys())
        
        def find_best_column(exact_terms, sub_terms):
            for k in header_options:
                col_title = k.split(":", 1)[1].strip() if ":" in k else k
                for term in exact_terms:
                    if col_title == term:
                        return k
            for k in header_options:
                col_title = k.split(":", 1)[1].strip() if ":" in k else k
                for term in sub_terms:
                    if term in col_title and "ส่วนต่าง" not in col_title and "จำนวน" not in col_title and "รับเงิน" not in col_title:
                        return k
            return header_options[0]

        default_date_up = find_best_column(["วัน/เดือน/ปี ขึ้นสินค้า", "วันที่ขึ้นสินค้า", "วันที่ขึ้น"], ["ขึ้น"])
        default_date_down = find_best_column(["วัน/เดือน/ปี ลงสินค้า", "วันที่ลงสินค้า", "วันที่ลง"], ["ลง"])
        default_plate = find_best_column(["ทะเบียน", "ทะเบียนรถ"], ["ทะเบียน"])
        default_prod = find_best_column(["สินค้า"], ["สินค้า"])
        default_origin = find_best_column(["ต้นทาง", "สถานที่ขึ้น"], ["ต้นทาง"])
        default_dest = find_best_column(["ปลายทาง", "สถานที่ลง"], ["ปลายทาง"])
        default_w_start = find_best_column(["นน.ต้นทาง", "น้ำหนักต้นทาง"], ["ต้นทาง"])
        default_w_end = find_best_column(["นน.ปลายทาง", "น้ำหนักปลายทาง"], ["ปลายทาง"])
        default_price = find_best_column(["ราคา"], ["ราคา", "บาท"])

        idx_plate = headers_found[default_plate]
        idx_prod = headers_found[default_prod]
        idx_date_up = headers_found[default_date_up]
        idx_date_down = headers_found[default_date_down]
        idx_origin = headers_found[default_origin]
        idx_dest = headers_found[default_dest]
        idx_w_start = headers_found[default_w_start]
        idx_w_end = headers_found[default_w_end]
        idx_price = headers_found[default_price]

        items_data = []
        last_valid_date = "-"
        
        for row in range(header_row + 1, ws.max_row + 1):
            plate = ws.cell(row=row, column=idx_plate).value
            plate_str = str(plate or "").strip()
            
            if plate_str and plate_str not in ["ทะเบียน", "None", "", "รวม", "ยอดรวม"]:
                raw_date_up = ws.cell(row=row, column=idx_date_up).value
                raw_date_down = ws.cell(row=row, column=idx_date_down).value
                
                formatted_d_up = format_date_thai(raw_date_up)
                formatted_d_down = format_date_thai(raw_date_down)
                
                if formatted_d_up and formatted_d_up != "-":
                    last_valid_date = formatted_d_up
                
                product = ws.cell(row=row, column=idx_prod).value or "ข้าวโพดแห้ง"
                origin = ws.cell(row=row, column=idx_origin).value or selected_display
                destination = ws.cell(row=row, column=idx_dest).value or "-"
                
                weight_start = ws.cell(row=row, column=idx_w_start).value or 0
                weight_end = ws.cell(row=row, column=idx_w_end).value or 0
                raw_price = ws.cell(row=row, column=idx_price).value or 0
                
                price_val = 0.62
                if isinstance(raw_price, (int, float)):
                    price_val = float(raw_price)
                elif isinstance(raw_price, str):
                    clean_p = raw_price.replace("/ตัน", "").replace("บาท", "").strip()
                    try:
                        price_val = float(clean_p)
                    except:
                        price_val = 0.62
                        
                if price_val > 100:
                    price_val = price_val / 1000.0

                items_data.append({
                    "date_up": last_valid_date,
                    "date_down": formatted_d_down,
                    "plate": plate_str,
                    "product": str(product).strip(),
                    "origin": str(origin).strip(),
                    "destination": str(destination).strip(),
                    "weight_start": weight_start,
                    "weight_end": weight_end,
                    "price_val": price_val
                })

        if len(items_data) == 0:
            st.warning(f"⚠️ ไม่พบข้อมูลรายการในชีท **{selected_display}**")
        else:
            unique_dates = list(set([item["date_up"] for item in items_data if item["date_up"] != "-"]))
            date_filter_options = ["แสดงทั้งหมด"] + unique_dates
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                selected_date = st.selectbox("📅 กรองตามวันที่ขึ้นสินค้า:", date_filter_options)
            
            if selected_date != "แสดงทั้งหมด":
                filtered_items = [item for item in items_data if item["date_up"] == selected_date]
            else:
                filtered_items = items_data

            with col_f2:
                st.metric(label="จำนวนรายการที่พบ", value=f"{len(filtered_items)} รายการ")
            
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
                    
                w_val = item['weight_end'] if item['weight_end'] != 0 else item['weight_start']
                w_str = f"{w_val:,.0f}" if isinstance(w_val, (int, float)) else f"{w_val}"
                
                label = f"ข้อ {idx} [{item['date_up']}]: ทะเบียน {item['plate']} | สินค้า: {item['product']} | ปลายทาง: {item['destination']} | นน.: {w_str}"
                
                is_selected = st.checkbox(label, key=key_name)
                if is_selected:
                    selected_indices.append(item)
            
            st.write("---")
            
            if st.button("🚀 สร้างไฟล์ Word ใบวางบิล (.docx)", type="primary", use_container_width=True):
                if len(selected_indices) == 0:
                    st.error("กรุณาเลือกอย่างน้อย 1 รายการก่อนสร้างเอกสาร")
                else:
                    doc = Document()
                    
                    section = doc.sections[0]
                    section.orientation = WD_ORIENTATION.LANDSCAPE
                    section.page_width = Inches(11.69)
                    section.page_height = Inches(8.27)
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                        
                    # 1. หัวเอกสาร
                    p_title = doc.add_paragraph()
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_title.paragraph_format.space_after = Pt(4)
                    r_title = p_title.add_run("ใบวางบิล")
                    r_title.font.name = "TH SarabunPSK"
                    r_title.font.size = Pt(22)
                    r_title.font.bold = True
                    
                    custom_date_be_str = format_date_thai(bill_date_input)
                    p_date = doc.add_paragraph()
                    p_date.paragraph_format.space_after = Pt(12)
                    r_date = p_date.add_run(f"วันที่ : {custom_date_be_str}")
                    r_date.font.name = "TH SarabunPSK"
                    r_date.font.size = Pt(18)
                    r_date.font.bold = True
                    
                    # 2. ตารางหลัก
                    total_rows = len(selected_indices) + 3
                    table = doc.add_table(rows=total_rows, cols=11)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    set_table_borders(table)
                    
                    headers = ["ลำดับ", "ทะเบียน", "สินค้า", "วันที่ขึ้น", "สถานที่ขึ้น", "สถานที่ลง", "วันที่ลง", "นน.\nต้นทาง", "นน.\nปลายทาง", "บาท/กก.", "ค่าขนส่ง"]
                    
                    hdr_cells = table.rows[0].cells
                    for i, head_text in enumerate(headers):
                        set_cell_background(hdr_cells[i], "4A607A")
                        p = hdr_cells[i].paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run(head_text)
                        r.font.name = "TH SarabunPSK"
                        r.font.size = Pt(14)
                        r.font.color.rgb = RGBColor(255, 255, 255)
                        
                    sum_w_start = 0
                    sum_w_end = 0
                    sum_total_shipping = 0
                    
                    for idx, item in enumerate(selected_indices, 1):
                        row_cells = table.rows[idx].cells
                        row_num = idx + 1
                        
                        w_start = item['weight_start'] if isinstance(item['weight_start'], (int, float)) else 0
                        w_end = item['weight_end'] if isinstance(item['weight_end'], (int, float)) else 0
                        price_kg = item['price_val']
                        
                        calc_w = w_end if w_end > 0 else w_start
                        shipping_cost = calc_w * price_kg
                        
                        sum_w_start += w_start
                        sum_w_end += w_end
                        sum_total_shipping += shipping_cost
                        
                        w_start_str = f"{w_start:,.0f}" if w_start > 0 else "-"
                        w_end_str = f"{w_end:,.0f}" if w_end > 0 else "-"
                        shipping_str = f"{shipping_cost:,.2f}" if shipping_cost > 0 else "-"
                        
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
                            f"{price_kg:.2f}"
                        ]
                        
                        for c_idx, val in enumerate(row_data):
                            p = row_cells[c_idx].paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c_idx in [7, 8, 9] else (WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 1, 3, 6] else WD_ALIGN_PARAGRAPH.LEFT)
                            r = p.add_run(val)
                            r.font.name = "TH SarabunPSK"
                            r.font.size = Pt(14)
                            
                        col_w_letter = 'I' if w_end > 0 else 'H'
                        word_formula = f"={col_w_letter}{row_num}*J{row_num}"
                        add_word_field_formula(row_cells[10], word_formula, shipping_str, is_bold=False, font_size=14)

                    # 3. แถวรวม
                    row_sum = table.rows[-2].cells
                    p_sum_lbl = row_sum[6].paragraphs[0]
                    p_sum_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r_sum_lbl = p_sum_lbl.add_run("รวม")
                    r_sum_lbl.font.name = "TH SarabunPSK"
                    r_sum_lbl.font.bold = True
                    r_sum_lbl.font.size = Pt(14)
                    
                    p_w_s = row_sum[7].paragraphs[0]
                    p_w_s.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r_w_s = p_w_s.add_run(f"{sum_w_start:,.0f}")
                    r_w_s.font.name = "TH SarabunPSK"
                    r_w_s.font.bold = True
                    r_w_s.font.size = Pt(14)
                    
                    p_w_e = row_sum[8].paragraphs[0]
                    p_w_e.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r_w_e = p_w_e.add_run(f"{sum_w_end:,.0f}" if sum_w_end > 0 else "0")
                    r_w_e.font.name = "TH SarabunPSK"
                    r_w_e.font.bold = True
                    r_w_e.font.size = Pt(14)
                    
                    add_word_field_formula(row_sum[10], "=SUM(ABOVE)", f"{sum_total_shipping:,.2f}", is_bold=True, font_size=14)

                    # 4. แถวยอดสุทธิ
                    row_net = table.rows[-1].cells
                    for cell in row_net:
                        set_cell_background(cell, "E8EEF8")
                        
                    row_net[0].merge(row_net[9])
                    p_net_lbl = row_net[0].paragraphs[0]
                    p_net_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r_net_lbl = p_net_lbl.add_run("ยอดสุทธิ   ")
                    r_net_lbl.font.name = "TH SarabunPSK"
                    r_net_lbl.font.bold = True
                    r_net_lbl.font.size = Pt(14)
                    
                    add_word_field_formula(row_net[10], "=K" + str(total_rows - 1), f"{sum_total_shipping:,.2f}", is_bold=True, font_size=14)

                    # 5. รายละเอียดการชำระเงิน (ตารางสั้นกระชับ กำหนดขนาดแน่นอน)
                    doc.add_paragraph().paragraph_format.space_after = Pt(12)
                    
                    p_pay_title = doc.add_paragraph()
                    p_pay_title.paragraph_format.space_after = Pt(4)
                    r_pay_title = p_pay_title.add_run("รายละเอียดการชำระเงิน")
                    r_pay_title.font.name = "TH SarabunPSK"
                    r_pay_title.font.size = Pt(18)
                    r_pay_title.font.bold = True
                    
                    pay_table = doc.add_table(rows=3, cols=2)
                    pay_table.alignment = WD_TABLE_ALIGNMENT.LEFT
                    set_table_borders(pay_table)
                    
                    pay_data = [
                        ("ชื่อบัญชี :", "หจก.ทีเอ็นพี โลจิสติกส์"),
                        ("ธนาคาร :", "กสิกรไทย"),
                        ("เลขที่บัญชี :", "097-1-01627-2")
                    ]
                    
                    # กำหนดความกว้างคอลัมน์ให้กระชับ
                    col_widths = [Inches(2.0), Inches(4.5)]
                    
                    for r_idx, (label, val) in enumerate(pay_data):
                        row_cells = pay_table.rows[r_idx].cells
                        
                        # คอลัมน์ซ้าย
                        row_cells[0].width = col_widths[0]
                        p_lbl = row_cells[0].paragraphs[0]
                        r_lbl = p_lbl.add_run(label)
                        r_lbl.font.name = "TH SarabunPSK"
                        r_lbl.font.size = Pt(16)
                        
                        # คอลัมน์ขวา
                        row_cells[1].width = col_widths[1]
                        p_val = row_cells[1].paragraphs[0]
                        r_val = p_val.add_run(val)
                        r_val.font.name = "TH SarabunPSK"
                        r_val.font.size = Pt(16)
                    
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