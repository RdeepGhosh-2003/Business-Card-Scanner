import http.server
import socketserver
import json
import smtplib
import sys
import os
import io
import subprocess
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

def ensure_docx():
    global docx, Pt, RGBColor, Inches, WD_ALIGN_PARAGRAPH, parse_xml, nsdecls, qn
    try:
        import docx as _docx
        from docx.shared import Pt as _Pt, RGBColor as _RGBColor, Inches as _Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml as _parse_xml
        from docx.oxml.ns import nsdecls as _nsdecls, qn as _qn
        docx, Pt, RGBColor, Inches = _docx, _Pt, _RGBColor, _Inches
        WD_ALIGN_PARAGRAPH, parse_xml, nsdecls, qn = _WD_ALIGN_PARAGRAPH, _parse_xml, _nsdecls, _qn
        return True
    except (ImportError, AttributeError):
        pass

    try:
        print("[*] Installing required library 'python-docx'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
        import docx as _docx
        from docx.shared import Pt as _Pt, RGBColor as _RGBColor, Inches as _Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml as _parse_xml
        from docx.oxml.ns import nsdecls as _nsdecls, qn as _qn
        docx, Pt, RGBColor, Inches = _docx, _Pt, _RGBColor, _Inches
        WD_ALIGN_PARAGRAPH, parse_xml, nsdecls, qn = _WD_ALIGN_PARAGRAPH, _parse_xml, _nsdecls, _qn
        return True
    except Exception as e:
        print(f"[!] Could not auto-install python-docx: {e}")
        docx = None
        return False

ensure_docx()

PORT = 8080
if len(sys.argv) > 1:
    try:
        PORT = int(sys.argv[1])
    except ValueError:
        pass

DIRECTORY = os.path.dirname(os.path.abspath(__file__))

def format_date_dmy(date_str):
    if not date_str:
        return ""
    parts = date_str.split('-')
    if len(parts) == 3 and len(parts[0]) == 4:
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return date_str.replace('-', '/')

def set_cell_border_none(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is not None:
        tcPr.remove(tcBorders)
    none_borders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                             '<w:top w:val="none"/>'
                             '<w:left w:val="none"/>'
                             '<w:bottom w:val="none"/>'
                             '<w:right w:val="none"/>'
                             '</w:tcBorders>')
    tcPr.append(none_borders)



def read_locked_file(path):
    import ctypes
    from ctypes import wintypes
    GENERIC_READ = 0x80000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80

    CreateFileW = ctypes.windll.kernel32.CreateFileW
    CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    CreateFileW.restype = wintypes.HANDLE

    ReadFile = ctypes.windll.kernel32.ReadFile
    ReadFile.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p]
    ReadFile.restype = wintypes.BOOL

    GetFileSize = ctypes.windll.kernel32.GetFileSize
    GetFileSize.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    GetFileSize.restype = wintypes.DWORD

    CloseHandle = ctypes.windll.kernel32.CloseHandle
    CloseHandle.argtypes = [wintypes.HANDLE]
    CloseHandle.restype = wintypes.BOOL

    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

    handle = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
    if handle == INVALID_HANDLE_VALUE:
        with open(path, 'rb') as f:
            return f.read()

    try:
        size = GetFileSize(handle, None)
        buf = ctypes.create_string_buffer(size)
        bytes_read = wintypes.DWORD()
        success = ReadFile(handle, buf, size, ctypes.byref(bytes_read), None)
        if not success:
            with open(path, 'rb') as f:
                return f.read()
        return buf.raw[:bytes_read.value]
    finally:
        CloseHandle(handle)


def fill_questionnaire_docx(card_data, output_path=None):
    """
    Fills the QUESTIONARE FORM 2025.docx template preserving its exact Word layout.
    """
    template_path = os.path.join(DIRECTORY, "QUESTIONARE FORM 2025.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    try:
        file_bytes = read_locked_file(template_path)
    except Exception:
        with open(template_path, 'rb') as f:
            file_bytes = f.read()

    if docx is None:
        if not ensure_docx():
            raise RuntimeError("The 'python-docx' library is required to generate Word documents. Please run: pip install python-docx")

    doc = docx.Document(io.BytesIO(file_bytes))

    # 1. Strip all checkbox drawings/picts from paragraph 1 onwards (Paragraph 0 has the Durga logo, KEEP it!)
    for p in doc.paragraphs[1:]:
        for d in p._p.xpath('.//w:drawing | .//w:pict'):
            d.getparent().remove(d)

    # 2. Clear drawings/pict from all table cells
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for d in c._tc.xpath('.//w:drawing | .//w:pict | .//w:tbl'):
                    d.getparent().remove(d)

    SECTION_INDENT = Inches(0.2)  # Matches template Interest in and Follow Up indent

    def add_p_bottom_border(p):
        pPr = p._p.get_or_add_pPr()
        pBdr = pPr.find(qn('w:pBdr'))
        if pBdr is not None:
            pPr.remove(pBdr)
        new_bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>')
        pPr.append(new_bdr)

    def set_label_and_val(p, label_text, val_text, placeholder_underline=""):
        p.text = ""
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0

        r_lbl = p.add_run(label_text)
        r_lbl.bold = True
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(13.5)
        r_lbl.font.color.rgb = RGBColor(0, 0, 0)

        if val_text:
            r_val = p.add_run(f"  {val_text}")
            r_val.bold = False
            r_val.font.name = 'Calibri'
            r_val.font.size = Pt(13)
            r_val.font.color.rgb = RGBColor(0, 0, 128)
        elif placeholder_underline:
            r_val = p.add_run(f"  {placeholder_underline}")
            r_val.bold = False
            r_val.font.name = 'Calibri'
            r_val.font.size = Pt(13)
            r_val.font.color.rgb = RGBColor(128, 128, 128)

    name         = card_data.get('name', '').strip()
    show_name    = (card_data.get('showName', '') or card_data.get('show_name', '') or 'IMTEX 2025').strip()
    company      = card_data.get('company', '').strip()
    designation  = card_data.get('designation', '').strip()
    address      = card_data.get('address', '').strip()
    city         = card_data.get('city', '').strip()
    state        = card_data.get('state', '').strip()
    pincode      = card_data.get('pincode', '').strip()
    phone        = (card_data.get('phone', '') or card_data.get('phone2', '')).strip()
    mobile       = card_data.get('mobile', '').strip()
    email        = card_data.get('email', '').strip()
    show_date    = format_date_dmy(card_data.get('showDate', '').strip())
    attended_by  = card_data.get('attendedBy', '').strip()
    cust_type    = card_data.get('customerType', '').strip()
    cust_product = card_data.get('customerProduct', '').strip()
    department   = card_data.get('department', '').strip()
    interests    = card_data.get('interestIn', [])
    action_req   = card_data.get('actionRequired', [])
    follow_up    = card_data.get('followUp', '').strip()
    catalogue    = card_data.get('catalogue', '').strip()
    remarks      = (card_data.get('notes', '') or card_data.get('remarks', '')).strip()

    if cust_type:
        if cust_type.lower().startswith('others'):
            other_detail = cust_type[6:].lstrip(': -').strip()
            cust_type_display = f"Others - {other_detail}" if other_detail else "Others - "
        else:
            cust_type_display = cust_type
    else:
        cust_type_display = ""

    # Map all interests dynamically
    interest_map = {}
    ordered_brand_keys = []
    for item in interests:
        if ':' in item:
            k, v = item.split(':', 1)
            k_clean = k.strip()
            v_clean = v.strip()
            interest_map[k_clean.upper()] = (k_clean, v_clean)
            if k_clean.upper() not in [x.upper() for x in ordered_brand_keys]:
                ordered_brand_keys.append(k_clean)
        else:
            k_clean = item.strip()
            interest_map[k_clean.upper()] = (k_clean, 'Yes')
            if k_clean.upper() not in [x.upper() for x in ordered_brand_keys]:
                ordered_brand_keys.append(k_clean)

    base_standard = ['INA', 'FAG', 'NSK', 'THK', 'TPI', 'BMD', 'SLF']
    all_brands_to_render = []
    for b in base_standard:
        all_brands_to_render.append(b)
    for b in ordered_brand_keys:
        if b.upper() not in [x.upper() for x in all_brands_to_render] and b.upper() != 'OTHERS':
            all_brands_to_render.append(b)
    all_brands_to_render.append('Others')

    selected_loc = state or city or ''
    loc_display = ""
    for loc in ['Mumbai', 'Chennai', 'Bengaluru']:
        if loc.lower() in selected_loc.lower():
            loc_display = loc
            break
    if not loc_display and selected_loc:
        loc_display = selected_loc

    # Table 0: Top Right Location
    if len(doc.tables) >= 1:
        t0 = doc.tables[0]
        tblPr = t0._tbl.tblPr
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        none_tbl_borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                                     '<w:top w:val="none"/>'
                                     '<w:left w:val="none"/>'
                                     '<w:bottom w:val="none"/>'
                                     '<w:right w:val="none"/>'
                                     '<w:insideH w:val="none"/>'
                                     '<w:insideV w:val="none"/>'
                                     '</w:tblBorders>')
        tblPr.append(none_tbl_borders)
        for r in t0.rows:
            for cell in r.cells:
                set_cell_border_none(cell)
                cell.text = ""
        if len(t0.rows) > 1 and len(t0.rows[1].cells) > 2:
            p_loc = t0.rows[1].cells[2].paragraphs[0]
            p_loc.text = loc_display
            p_loc.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if len(p_loc.runs) > 0:
                p_loc.runs[0].bold = True
                p_loc.runs[0].font.size = Pt(16)
                p_loc.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    # Table 1: Left Table (Visitor Info)
    # Width: 5200 dxa (~3.61 in)
    # Position: tblpX = 198 dxa from margin (Left edge)
    if len(doc.tables) >= 2:
        t1 = doc.tables[1]
        tblpPr1 = t1._tbl.xpath('.//w:tblpPr')
        if tblpPr1:
            tblpPr1[0].set(qn('w:tblpX'), '198')
            tblpPr1[0].set(qn('w:horzAnchor'), 'margin')

        for r in t1.rows:
            trH = r._tr.xpath('.//w:trHeight')
            if trH:
                r._tr.trPr.remove(trH[0])
            for c in r.cells:
                tcW = c._tc.xpath('.//w:tcW')
                if tcW:
                    tcW[0].set(qn('w:w'), '5200')
                    tcW[0].set(qn('w:type'), 'dxa')

        if len(t1.rows) > 0:
            set_label_and_val(t1.rows[0].cells[0].paragraphs[0], 'Name of the Show: ', show_name)
        if len(t1.rows) > 1:
            set_label_and_val(t1.rows[1].cells[0].paragraphs[0], 'Visitor Name: ', name)
        if len(t1.rows) > 2:
            set_label_and_val(t1.rows[2].cells[0].paragraphs[0], 'Company Name: ', company)
        if len(t1.rows) > 3:
            set_label_and_val(t1.rows[3].cells[0].paragraphs[0], 'Designation: ', designation)
        if len(t1.rows) > 4:
            cell_addr = t1.rows[4].cells[0]
            set_label_and_val(cell_addr.paragraphs[0], 'Address: ', address)
            while len(cell_addr.paragraphs) > 1:
                p_ex = cell_addr.paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)
        if len(t1.rows) > 5:
            p = t1.rows[5].cells[0].paragraphs[0]
            p.text = ""
            p.paragraph_format.space_before = Pt(0.5)
            p.paragraph_format.space_after = Pt(0.5)
            r_c_lbl = p.add_run('City: ')
            r_c_lbl.bold = True
            r_c_lbl.font.size = Pt(13.5)
            city_part = city if city else (state if state else "")
            r_c_val = p.add_run(f"  {city_part}" if city_part else "  __________")
            r_c_val.bold = False
            r_c_val.font.size = Pt(13)
            r_c_val.font.color.rgb = RGBColor(0, 0, 128) if city_part else RGBColor(128, 128, 128)

            r_p_lbl = p.add_run('   Pin Code: ')
            r_p_lbl.bold = True
            r_p_lbl.font.size = Pt(13.5)
            r_p_val = p.add_run(f"  {pincode}" if pincode else "  ______")
            r_p_val.bold = False
            r_p_val.font.size = Pt(13)
            r_p_val.font.color.rgb = RGBColor(0, 0, 128) if pincode else RGBColor(128, 128, 128)
        if len(t1.rows) > 6:
            set_label_and_val(t1.rows[6].cells[0].paragraphs[0], 'Phone: ', phone)
        if len(t1.rows) > 7:
            set_label_and_val(t1.rows[7].cells[0].paragraphs[0], 'Mobile No: ', mobile)
        if len(t1.rows) > 8:
            set_label_and_val(t1.rows[8].cells[0].paragraphs[0], 'Email: ', email)

    # Table 2: Right Table (Show Details)
    # Width: 4600 dxa (~3.19 in)
    # Position: tblpX = 6200 dxa from page (strictly right of Table 1)
    if len(doc.tables) >= 3:
        t2 = doc.tables[2]
        tblpPr2 = t2._tbl.xpath('.//w:tblpPr')
        if tblpPr2:
            tblpPr2[0].set(qn('w:tblpX'), '6200')
            tblpPr2[0].set(qn('w:horzAnchor'), 'page')

        for r in t2.rows:
            trH = r._tr.xpath('.//w:trHeight')
            if trH:
                r._tr.trPr.remove(trH[0])
            for c in r.cells:
                tcW = c._tc.xpath('.//w:tcW')
                if tcW:
                    tcW[0].set(qn('w:w'), '4600')
                    tcW[0].set(qn('w:type'), 'dxa')

        if len(t2.rows) > 0:
            p0 = t2.rows[0].cells[0].paragraphs[0]
            p0.text = show_name
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p0.paragraph_format.space_before = Pt(0.5)
            p0.paragraph_format.space_after = Pt(0.5)
            if len(p0.runs) > 0:
                p0.runs[0].bold = True
                p0.runs[0].font.size = Pt(14.5)
            while len(t2.rows[0].cells[0].paragraphs) > 1:
                p_ex = t2.rows[0].cells[0].paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)
        if len(t2.rows) > 1:
            cell1 = t2.rows[1].cells[0]
            set_label_and_val(cell1.paragraphs[0], 'Day of the show: ', show_date, '____________________')
            while len(cell1.paragraphs) > 1:
                p_ex = cell1.paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)
        if len(t2.rows) > 2:
            cell2 = t2.rows[2].cells[0]
            set_label_and_val(cell2.paragraphs[0], 'Customer Type: ', cust_type_display, '____________________')
            while len(cell2.paragraphs) > 1:
                p_ex = cell2.paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)
        if len(t2.rows) > 3:
            cell3 = t2.rows[3].cells[0]
            cell3.paragraphs[0].text = ""
            cell3.paragraphs[0].paragraph_format.space_before = Pt(0.5)
            cell3.paragraphs[0].paragraph_format.space_after = Pt(0.5)
            r_cp = cell3.paragraphs[0].add_run("Customer's Product:")
            r_cp.bold = True
            r_cp.font.size = Pt(13.5)
            lines = [l.strip() for l in cust_product.split('\n') if l.strip()] if cust_product else []
            for i in range(max(len(lines), 2)):
                line_text = lines[i] if i < len(lines) else ""
                if i + 1 < len(cell3.paragraphs):
                    p_line = cell3.paragraphs[i + 1]
                    p_line.text = ""
                else:
                    p_line = cell3.add_paragraph()
                p_line.paragraph_format.space_before = Pt(0)
                p_line.paragraph_format.space_after = Pt(0)
                p_line.paragraph_format.line_spacing = 1.0
                if line_text:
                    r = p_line.add_run(line_text)
                    r.bold = False
                    r.font.size = Pt(12.5)
                    r.font.color.rgb = RGBColor(0, 0, 128)
                else:
                    p_line.text = "__________________________________________"
            while len(cell3.paragraphs) > max(len(lines) + 1, 3):
                p_ex = cell3.paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)
        if len(t2.rows) > 4:
            cell4 = t2.rows[4].cells[0]
            set_label_and_val(cell4.paragraphs[0], 'Department: ', department, '____________________')
            while len(cell4.paragraphs) > 1:
                p_ex = cell4.paragraphs[-1]._p
                p_ex.getparent().remove(p_ex)

    # ── Spacer Paragraphs p[01]..p[05] (Canvas height for native floating tables) ──
    addr_line_count = len(address.split('\n')) if address else 1
    prod_line_count = len([l for l in (cust_product or '').split('\n') if l.strip()])
    extra_addr_lines = max(0, addr_line_count - 1)
    extra_prod_lines = max(0, prod_line_count - 2)

    extra_needed = max(extra_addr_lines * 6.5, max(0.0, extra_prod_lines * 6.5 - 15.0))

    for i in range(1, 6):
        sp_p = doc.paragraphs[i]
        sp_p.text = ""
        sp_p.paragraph_format.space_before = Pt(0)
        sp_p.paragraph_format.space_after = Pt(0)
        sp_p.paragraph_format.line_spacing = Pt(14)

    doc.paragraphs[5].paragraph_format.space_after = Pt(extra_needed)

    # ── Paragraph 6: Interest in Heading (Strictly Left Aligned) ──
    p6 = doc.paragraphs[6]
    p6.text = "Interest in"
    p6.runs[0].bold = True
    p6.runs[0].font.size = Pt(14.5)
    p6.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p6.paragraph_format.space_before = Pt(1)
    p6.paragraph_format.space_after = Pt(0.5)
    p6.paragraph_format.left_indent = SECTION_INDENT
    p6.paragraph_format.first_line_indent = Inches(0)

    # Find p_div1 (the divider border line paragraph between Interest in and Action Required)
    p_div1 = doc.paragraphs[23]

    # Remove all empty checkbox paragraphs between p6 and p_div1
    p_current = p6._p.getnext()
    while p_current is not None and p_current != p_div1._p:
        p_next = p_current.getnext()
        p_current.getparent().remove(p_current)
        p_current = p_next

    brand_font_size = Pt(13) if len(all_brands_to_render) <= 8 else Pt(11.5)

    # Insert all brands sequentially before p_div1
    for b_name in all_brands_to_render:
        new_p_elm = parse_xml(f'<w:p {nsdecls("w")}/>')
        p_div1._p.addprevious(new_p_elm)
        p_obj = docx.text.paragraph.Paragraph(new_p_elm, doc)

        p_obj.text = ""
        p_obj.paragraph_format.space_before = Pt(0)
        p_obj.paragraph_format.space_after = Pt(0)
        p_obj.paragraph_format.line_spacing = 1.0
        p_obj.paragraph_format.left_indent = SECTION_INDENT
        p_obj.paragraph_format.first_line_indent = Inches(0)
        p_obj.alignment = WD_ALIGN_PARAGRAPH.LEFT

        r_b = p_obj.add_run(f"{b_name}:")
        r_b.bold = True
        r_b.font.size = brand_font_size

        note_tuple = interest_map.get(b_name.upper(), ('', ''))
        note = note_tuple[1] if isinstance(note_tuple, tuple) else note_tuple
        if note and note != 'Yes':
            r_v = p_obj.add_run(f"  {note}")
            r_v.bold = False
            r_v.font.size = brand_font_size
            r_v.font.color.rgb = RGBColor(0, 0, 128)
        elif note == 'Yes':
            r_v = p_obj.add_run(f"  Yes")
            r_v.bold = False
            r_v.font.size = brand_font_size
            r_v.font.color.rgb = RGBColor(0, 0, 128)
        else:
            r_v = p_obj.add_run("  ______________________________________________________________________________")
            r_v.bold = False
            r_v.font.size = brand_font_size
            r_v.font.color.rgb = RGBColor(128, 128, 128)

    # Keep p_div1 as the bottom border line of the Interest in section
    p_div1.text = ""
    p_div1.paragraph_format.space_before = Pt(0)
    p_div1.paragraph_format.space_after = Pt(0)
    p_div1.paragraph_format.line_spacing = 1.0

    # 5. Action Required (Paragraph following p_div1 with bottom border)
    p_act = p_div1._p.getnext()
    p_act_obj = docx.text.paragraph.Paragraph(p_act, doc)
    p_act_obj.text = ""
    p_act_obj.paragraph_format.space_before = Pt(1)
    p_act_obj.paragraph_format.space_after = Pt(2)
    p_act_obj.paragraph_format.line_spacing = 1.0
    p_act_obj.paragraph_format.left_indent = SECTION_INDENT
    p_act_obj.paragraph_format.first_line_indent = Inches(0)
    p_act_obj.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_act_lbl = p_act_obj.add_run("Action Required: ")
    r_act_lbl.bold = True
    r_act_lbl.font.size = Pt(13.5)
    if action_req:
        r_act_v = p_act_obj.add_run("   |   ".join(action_req))
        r_act_v.bold = False
        r_act_v.font.size = Pt(13)
        r_act_v.font.color.rgb = RGBColor(0, 0, 128)
    else:
        r_act_v = p_act_obj.add_run("_____________________________________________________________________")
        r_act_v.bold = False
        r_act_v.font.size = Pt(13)
        r_act_v.font.color.rgb = RGBColor(128, 128, 128)

    # Insert full-width divider border line right below Action Required (no left indent)
    new_div_act = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>')
    p_act.addnext(new_div_act)

    # Find Remarks
    p_rem = None
    for p in doc.paragraphs:
        if "Remarks" in p.text:
            p_rem = p
            break

    # Remove all paragraphs between new_div_act and p_rem (removes p_div2, p_fu, p_div3)
    p_curr = new_div_act.getnext()
    while p_curr is not None and p_curr != p_rem._p:
        p_n = p_curr.getnext()
        p_curr.getparent().remove(p_curr)
        p_curr = p_n

    # 6. Follow Up & Product Catalogue (2-column layout with direct bottom border, zero dead space!)
    tbl_fu = doc.add_table(rows=1, cols=2)
    p_rem._p.addprevious(tbl_fu._tbl)

    tblPr = tbl_fu._tbl.tblPr
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    tblPr.append(parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                           '<w:top w:val="none"/>'
                           '<w:left w:val="none"/>'
                           '<w:bottom w:val="none"/>'
                           '<w:right w:val="none"/>'
                           '<w:insideH w:val="none"/>'
                           '<w:insideV w:val="none"/>'
                           '</w:tblBorders>'))
    tbl_fu.columns[0].width = Inches(2.2)
    tbl_fu.columns[1].width = Inches(4.8)

    cell_fu = tbl_fu.cell(0, 0)
    cell_cat = tbl_fu.cell(0, 1)
    set_cell_border_none(cell_fu)
    set_cell_border_none(cell_cat)

    p_c_fu = cell_fu.paragraphs[0]
    p_c_fu.text = ""
    p_c_fu.paragraph_format.space_before = Pt(1)
    p_c_fu.paragraph_format.space_after = Pt(2)
    p_c_fu.paragraph_format.line_spacing = 1.0
    p_c_fu.paragraph_format.left_indent = SECTION_INDENT
    r_fu_lbl = p_c_fu.add_run("Follow Up: ")
    r_fu_lbl.bold = True
    r_fu_lbl.font.size = Pt(13.5)
    if follow_up:
        r_fu_v = p_c_fu.add_run(follow_up)
        r_fu_v.bold = False
        r_fu_v.font.size = Pt(13)
        r_fu_v.font.color.rgb = RGBColor(0, 0, 128)
    else:
        r_fu_v = p_c_fu.add_run("__________")
        r_fu_v.bold = False
        r_fu_v.font.size = Pt(13)
        r_fu_v.font.color.rgb = RGBColor(128, 128, 128)

    p_c_cat = cell_cat.paragraphs[0]
    p_c_cat.text = ""
    p_c_cat.paragraph_format.space_before = Pt(1)
    p_c_cat.paragraph_format.space_after = Pt(2)
    p_c_cat.paragraph_format.line_spacing = 1.0
    r_cat_lbl = p_c_cat.add_run("Product Catalogue Issued: ")
    r_cat_lbl.bold = True
    r_cat_lbl.font.size = Pt(13.5)
    if catalogue:
        r_cat_v = p_c_cat.add_run(catalogue)
        r_cat_v.bold = False
        r_cat_v.font.size = Pt(13)
        r_cat_v.font.color.rgb = RGBColor(0, 0, 128)
    else:
        r_cat_v = p_c_cat.add_run("____________________")
        r_cat_v.bold = False
        r_cat_v.font.size = Pt(13)
        r_cat_v.font.color.rgb = RGBColor(128, 128, 128)

    # Continuous smooth full-width divider line between Follow Up and Remarks
    new_div_fu = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>')
    p_rem._p.addprevious(new_div_fu)

    # 7. Remarks
    p_rem.text = ""
    p_rem.paragraph_format.space_before = Pt(1)
    p_rem.paragraph_format.space_after = Pt(0)
    p_rem.paragraph_format.left_indent = SECTION_INDENT
    p_rem.paragraph_format.first_line_indent = Inches(0)
    p_rem.alignment = WD_ALIGN_PARAGRAPH.LEFT

    r_rem = p_rem.add_run("Remarks: ")
    r_rem.bold = True
    r_rem.font.size = Pt(13.5)
    curr_p = p_rem
    if remarks:
        rem_lines = [l.strip() for l in remarks.split('\n') if l.strip()]
        r_rem_v = p_rem.add_run(rem_lines[0])
        r_rem_v.bold = False
        r_rem_v.font.size = Pt(12.5)
        r_rem_v.font.color.rgb = RGBColor(0, 0, 128)

        for rem_line in rem_lines[1:]:
            new_p_elm = parse_xml(f'<w:p {nsdecls("w")}/>')
            curr_p._p.addnext(new_p_elm)
            p_next = docx.text.paragraph.Paragraph(new_p_elm, doc)
            p_next.paragraph_format.space_before = Pt(0.5)
            p_next.paragraph_format.space_after = Pt(0)
            p_next.paragraph_format.left_indent = SECTION_INDENT
            p_next.paragraph_format.first_line_indent = Inches(0)
            p_next.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r_next = p_next.add_run(rem_line)
            r_next.bold = False
            r_next.font.size = Pt(12.5)
            r_next.font.color.rgb = RGBColor(0, 0, 128)
            curr_p = p_next

    # 8. Attended By (Divider line above Attended by, exactly matching Remarks divider)
    p_att = None
    for p in doc.paragraphs:
        if "Attended by" in p.text:
            p_att = p
            break

    # Delete everything between curr_p and p_att!
    p_curr = curr_p._p.getnext()
    while p_curr is not None and p_curr != p_att._p:
        p_n = p_curr.getnext()
        p_curr.getparent().remove(p_curr)
        p_curr = p_n

    # Insert a clean divider border line above Attended By
    new_div_att = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>')
    p_att._p.addprevious(new_div_att)

    if p_att is not None:
        p_att.text = ""
        p_att.paragraph_format.space_before = Pt(1)
        p_att.paragraph_format.space_after = Pt(0)
        p_att.paragraph_format.left_indent = SECTION_INDENT
        p_att.paragraph_format.first_line_indent = Inches(0)
        p_att.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_att = p_att.add_run("Attended by: ")
        r_att.bold = True
        r_att.font.size = Pt(13.5)
        if attended_by:
            r_att_v = p_att.add_run(f"  {attended_by}")
            r_att_v.bold = False
            r_att_v.font.size = Pt(13)
            r_att_v.font.color.rgb = RGBColor(0, 0, 128)
        else:
            r_att_v = p_att.add_run("_____________________________________________________________________________")
            r_att_v.bold = False
            r_att_v.font.size = Pt(13)
            r_att_v.font.color.rgb = RGBColor(128, 128, 128)

    if output_path:
        doc.save(output_path)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

class AppHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # Enable CORS and avoid stale caching for local development
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/send-email':
            self.handle_send_email()
        elif self.path == '/api/test-smtp':
            self.handle_test_smtp()
        elif self.path == '/api/open-outlook-desktop':
            self.handle_open_outlook_desktop()
        elif self.path == '/api/fill-questionnaire':
            self.handle_fill_questionnaire()
        else:
            self.send_error(404, "Endpoint not found")

    def handle_send_email(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))

            to_addr = data.get('to', '').strip()
            subject = data.get('subject', '').strip()
            body = data.get('body', '').strip()
            footer = data.get('footer', '').strip()
            smtp_user = data.get('smtpUser', '').strip()
            smtp_pass = data.get('smtpPass', '').strip().replace(' ', '')
            smtp_host = data.get('smtpHost', 'smtp.gmail.com').strip()
            smtp_port = int(data.get('smtpPort', 587))
            sender_name = data.get('senderName', '').strip() or smtp_user

            # Determine provider name for friendly messaging
            provider = "Email Server"
            if "office365" in smtp_host.lower() or "outlook" in smtp_host.lower() or "hotmail" in smtp_host.lower():
                provider = "Outlook / Microsoft 365"
            elif "gmail" in smtp_host.lower():
                provider = "Gmail"

            if not to_addr:
                self.send_json_response(400, {'success': False, 'error': 'Recipient email address is required.'})
                return

            if not smtp_user or not smtp_pass:
                self.send_json_response(400, {
                    'success': False,
                    'error': f'{provider} credentials not configured. Please enter your email and password in Settings -> Direct Send.'
                })
                return

            # Combine body and footer
            full_body = body
            if footer:
                full_body = f"{body}\n\n--\n{footer}" if body else footer

            # Build MIME message
            msg = MIMEMultipart()
            msg['From'] = f"{sender_name} <{smtp_user}>" if sender_name and sender_name != smtp_user else smtp_user
            msg['To'] = to_addr
            msg['Subject'] = Header(subject, 'utf-8').encode()
            msg.attach(MIMEText(full_body, 'plain', 'utf-8'))

            # Send via SMTP
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=15)
            server.ehlo()
            if smtp_port == 587:
                server.starttls()
                server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, [to_addr], msg.as_string())
            server.quit()

            self.send_json_response(200, {
                'success': True,
                'message': f'Email successfully sent to {to_addr} directly via {provider}!'
            })
        except smtplib.SMTPAuthenticationError as e:
            provider = "Outlook / Microsoft 365" if ("office365" in smtp_host.lower() or "outlook" in smtp_host.lower()) else "Gmail"
            err_str = str(e)
            if "5.7.139" in err_str or "SmtpClientAuthentication" in err_str:
                auth_msg = 'Microsoft 365 Error: "Authenticated SMTP" is disabled for this mailbox by your organization. Enable it in Microsoft 365 Admin Center > Users > Mail > Manage email apps > Authenticated SMTP, or use the "Outlook Desktop App" button in the email modal.'
            elif provider == "Gmail":
                auth_msg = 'Gmail authentication failed. If 2-Step Verification is active, generate a 16-character App Password at https://myaccount.google.com/apppasswords.'
            else:
                auth_msg = 'Authentication failed. Please verify your email & password (or generate an App Password in Microsoft Security).'
            self.send_json_response(401, {
                'success': False,
                'error': auth_msg
            })
        except Exception as e:
            self.send_json_response(500, {
                'success': False,
                'error': f'Failed to send email: {str(e)}'
            })

    def handle_test_smtp(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))

            smtp_user = data.get('smtpUser', '').strip()
            smtp_pass = data.get('smtpPass', '').strip().replace(' ', '')
            smtp_host = data.get('smtpHost', 'smtp.gmail.com').strip()
            smtp_port = int(data.get('smtpPort', 587))

            provider = "Email Server"
            if "office365" in smtp_host.lower() or "outlook" in smtp_host.lower() or "hotmail" in smtp_host.lower():
                provider = "Outlook / Microsoft 365"
            elif "gmail" in smtp_host.lower():
                provider = "Gmail"

            if not smtp_user or not smtp_pass:
                self.send_json_response(400, {'success': False, 'error': f'Please provide both {provider} email address and password.'})
                return

            server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
            server.ehlo()
            if smtp_port == 587:
                server.starttls()
                server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.quit()

            self.send_json_response(200, {
                'success': True,
                'message': f'Connection successful! Verified {provider} login for {smtp_user}.'
            })
        except smtplib.SMTPAuthenticationError as e:
            provider = "Outlook / Microsoft 365" if ("office365" in smtp_host.lower() or "outlook" in smtp_host.lower()) else "Gmail"
            err_str = str(e)
            if "5.7.139" in err_str or "SmtpClientAuthentication" in err_str:
                auth_msg = 'Microsoft 365 Error 5.7.139: "Authenticated SMTP" is disabled for this mailbox by your organization. Enable it in Microsoft 365 Admin Center (Users > Mail > Manage email apps > Authenticated SMTP), or use the "Outlook Desktop App" button.'
            elif provider == "Gmail":
                auth_msg = 'Authentication failed. Please verify your Gmail address and 16-character App Password.'
            else:
                auth_msg = 'Authentication failed. Please verify your Outlook/Microsoft 365 email and password.'
            self.send_json_response(401, {
                'success': False,
                'error': auth_msg
            })
        except Exception as e:
            self.send_json_response(500, {
                'success': False,
                'error': f'Connection error: {str(e)}'
            })

    def handle_open_outlook_desktop(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))

            to = data.get('to', '').strip()
            subject = data.get('subject', '').strip()
            body = data.get('body', '').strip()
            footer = data.get('footer', '').strip()
            full_body = f"{body}\n\n{footer}".strip() if footer else body

            params = []
            if subject:
                params.append(f"subject={urllib.parse.quote(subject)}")
            if full_body:
                params.append(f"body={urllib.parse.quote(full_body)}")
            param_str = "&".join(params)
            mailto = f"{to}?{param_str}" if param_str else to

            outlook_paths = [
                r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
            ]
            outlook_exe = next((p for p in outlook_paths if os.path.exists(p)), None)

            if outlook_exe:
                cmd = f'"{outlook_exe}" /c ipm.note /m "{mailto}"'
                subprocess.Popen(cmd, shell=True)
            else:
                os.startfile(f"mailto:{mailto}")

            self.send_json_response(200, {
                'success': True,
                'message': 'Opened in Microsoft Outlook Desktop App'
            })
        except Exception as e:
            self.send_json_response(500, {
                'success': False,
                'error': f'Failed to launch Outlook Desktop: {str(e)}'
            })

    def handle_fill_questionnaire(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            card_data = json.loads(post_data.decode('utf-8'))

            docx_bytes = fill_questionnaire_docx(card_data)

            visitor_name = card_data.get('name', '').strip() or 'Visitor'
            company = card_data.get('company', '').strip() or 'Company'
            safe_filename = f"Questionnaire_{visitor_name}_{company}.docx".replace(' ', '_').replace('/', '_').replace('\\', '_')

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{safe_filename}"')
            self.send_header('Content-Length', str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)
        except Exception as e:
            self.send_json_response(500, {
                'success': False,
                'error': f'Failed to generate questionnaire Word doc: {str(e)}'
            })

    def send_json_response(self, code, data):
        response_bytes = json.dumps(data).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(response_bytes)))
        self.end_headers()
        self.wfile.write(response_bytes)

def run():
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), AppHandler) as httpd:
        print(f"====================================================")
        print(f" Business Card Scanner Server running at:")
        print(f" http://localhost:{PORT}")
        print(f" Direct Gmail sending enabled (/api/send-email)")
        print(f"====================================================")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server.")
            httpd.server_close()

if __name__ == '__main__':
    run()
