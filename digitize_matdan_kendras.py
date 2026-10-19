import os
import sys
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.append(os.path.abspath("."))
import pdfplumber
from kruti_to_unicode import KrutiDev_to_Unicode

def clean_hindi(text):
    if not text:
        return ""
    s = KrutiDev_to_Unicode(str(text))
    replacements = [
        ("पदाध्िाकारी", "पदाधिकारी"),
        ("अध्न्र्ि ा", "अग्नि"),
        ("अध्िाक", "अधिक"),
        ("अध्िीक्षक", "अधीक्षक"),
        ("अन्त् ा", "अन्त"),
        ("अन्तर्ग्गत", "अन्तर्गत"),
        ("अन्तर्गत", "अन्तर्गत"),
        ("अंतगर्त", "अंतर्गत"),
        ("अवस्िथत", "अवस्थित"),
        ("उत्क्रम्िात", "उत्क्रमित"),
        ("उच्च् ", "उच्च "),
        ("उच्च्", "उच्च"),
        ("कार्र्यालय", "कार्यालय"),
        ("कृष्िा", "कृषि"),
        ("खगडि़या", "खगड़िया"),
        ("चकार्इ", "चकाई"),
        ("जमुर्इ", "जमुई"),
        ("दक्ष्िाण", "दक्षिण"),
        ("पख््र ांड", "प्रखंड"),
        ("पख््र", "प्रखंड"),
        ("पस््र तावित", "प्रस्तावित"),
        ("पस््र ताव", "प्रस्ताव"),
        ("पस््र", "प्र"),
        ("पश्िचमी", "पश्चिमी"),
        ("पा्रणपुर", "प्राणपुर"),
        ("पूण्र्िाया", "पूर्णिया"),
        ("पूिर्णया़", "पूर्णिया"),
        ("पूिर्णया", "पूर्णिया"),
        ("पवूर्", "पूर्व"),
        ("पूर्र् व", "पूर्व"),
        ("पूवर्", "पूर्व"),
        ("पूर्व्व त~", "पूर्ववत"),
        ("पूर्व्वत~", "पूर्ववत"),
        ("पूर्व्व त", "पूर्ववत"),
        ("पूर्व्वत", "पूर्ववत"),
        ("पूवर्व त्", "पूर्ववत"),
        ("पूवर्व त", "पूर्ववत"),
        ("पूवर्वत्", "पूर्ववत"),
        ("पूवर्वत", "पूर्ववत"),
        ("प्राथ्िामिक", "प्राथमिक"),
        ("प्रस्तार्िवत", "प्रस्तावित"),
        ("प्रर्िवष्ट", "प्रविष्ट"),
        ("बख्ितयारपुर", "बख्तियारपुर"),
        ("बले हर", "बेलहर"),
        ("बले दौर", "बेलदौर"),
        ("मध्य् ", "मध्य "),
        ("मध्य्", "मध्य"),
        ("माध्य्िामिक", "माध्यमिक"),
        ("विध्य् ा", "विद्या"),
        ("विद्य् ा", "विद्या"),
        ("वश्े म", "वेश्म"),
        ("वश्े", "वेश्म"),
        ("शेखपरु ा", "शेखपुरा"),
        ("मधेपरु ा", "मधेपुरा"),
        ("सम्पण्ूर् ा", "सम्पूर्ण"),
        ("सम्पूर्ण् ा", "सम्पूर्ण"),
        ("संपण्ूर्", "सम्पूर्ण"),
        ("साम्िुदायिक", "सामुदायिक"),
        ("साध्िारण", "साधारण"),
        ("सिमरी बख्ितयारपुर", "सिमरी बख्तियारपुर"),
        ("सोनवष्र्ाा", "सोनवर्षा"),
        ("स्िथत", "स्थित"),
        ("त्रिवण्े ाीगजं", "त्रिवेणीगंज"),
        ("त्रिवण्े", "त्रिवेणी"),
        ("प्रर् ा", "प्र"),
        ("र् ा", "र्"),
        ("य् ा", "य"),
        ("भ् ा", "भ"),
        ("ध्िा", "धि"),
        ("ध्ि", "धि"),
        ("न्र् ा", "न्र"),
        ("वन्द् ा", "वन्द"),
        ("एव ं", "एवं"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    lines = [re.sub(r'[ \t]+', ' ', line).strip() for line in s.split('\n')]
    return '\n'.join([l for l in lines if l])

def parse_voter_count(text):
    if not text:
        return 0
    m = re.search(r'कुल[- :]*([0-9]+)', text)
    if m:
        return int(m.group(1))
    m2 = re.search(r'^\s*([0-9]+)\s*$', text)
    if m2:
        return int(m2.group(1))
    numbers = [int(n) for n in re.findall(r'[0-9]+', text)]
    if len(numbers) == 1:
        return numbers[0]
    return numbers[-1] if numbers else 0

canonical_districts = [
    "खगड़िया", "सहरसा", "मधेपुरा", "सुपौल", "अररिया",
    "किशनगंज", "पूर्णिया", "कटिहार", "भागलपुर", "मुंगेर",
    "लखीसराय", "शेखपुरा", "जमुई", "बाँका"
]

def detect_district(row_text):
    for d in canonical_districts:
        if d == "जमुई" and ("जमुई" in row_text or "जमुर्इ" in row_text):
            return "जमुई"
        elif d == "पूर्णिया" and ("पूर्णिया" in row_text or "पूिर्णया" in row_text or "पूण्र्िाया" in row_text):
            return "पूर्णिया"
        elif d == "मधेपुरा" and ("मधेपुरा" in row_text or "मधेपरु ा" in row_text):
            return "मधेपुरा"
        elif d == "खगड़िया" and ("खगड़िया" in row_text or "खगडि़या" in row_text):
            return "खगड़िया"
        elif d in row_text:
            return d
    return "अज्ञात"

print("Extracting tables from PDF...")
raw_records = []
with pdfplumber.open("Matdan Kendron Ki Suchi.pdf") as pdf:
    for page_idx, page in enumerate(pdf.pages):
        tbl = page.extract_tables()[0]
        for row_idx, row in enumerate(tbl):
            first_c = clean_hindi(row[0]).strip()
            # If row starts with digit, it's a polling station
            if any(ch.isdigit() for ch in first_c):
                cleaned_cells = [clean_hindi(c) for c in row]
                # Pad to 10 cols if needed
                while len(cleaned_cells) < 10:
                    cleaned_cells.append("")
                row_full_text = " ".join(cleaned_cells)
                dist = detect_district(row_full_text)
                raw_records.append({
                    'page': page_idx + 1,
                    'row_idx': row_idx,
                    'serial': cleaned_cells[0],
                    'district': dist,
                    'location': cleaned_cells[1],
                    'building': cleaned_cells[2],
                    'area': cleaned_cells[3],
                    'gate': cleaned_cells[4],
                    'voter_area': cleaned_cells[5],
                    'voter_type': cleaned_cells[6],
                    'voters_assigned_str': cleaned_cells[7],
                    'voters_count': parse_voter_count(cleaned_cells[7]),
                    'max_distance': cleaned_cells[8],
                    'remarks': cleaned_cells[9]
                })

print(f"Total polling stations extracted: {len(raw_records)}")

# Create Excel Workbook
wb = openpyxl.Workbook()

# Style definitions
font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
font_subtitle = Font(name="Calibri", size=11, bold=True, color="1B365D")
font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_banner = Font(name="Calibri", size=12, bold=True, color="0F2942")
font_data = Font(name="Calibri", size=10, color="000000")
font_data_bold = Font(name="Calibri", size=10, bold=True, color="000000")
font_total = Font(name="Calibri", size=11, bold=True, color="000000")

fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_subtitle = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
fill_header = PatternFill(start_color="244061", end_color="244061", fill_type="solid")
fill_banner = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")
fill_row_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_row_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_total = PatternFill(start_color="EAEEF3", end_color="EAEEF3", fill_type="solid")

thin_border_side = Side(border_style="thin", color="CCCCCC")
med_border_top = Side(border_style="medium", color="1B365D")
double_border_bottom = Side(border_style="double", color="1B365D")

border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
border_banner = Border(left=thin_border_side, right=thin_border_side, top=Side(border_style="medium", color="244061"), bottom=Side(border_style="medium", color="244061"))
border_total = Border(top=med_border_top, bottom=double_border_bottom, left=thin_border_side, right=thin_border_side)

align_center_middle = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
align_center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
align_right_top = Alignment(horizontal="right", vertical="top", wrap_text=True)

# -------------------------------------------------------------
# SHEET 1: मतदान केन्द्रों की सूची (Exact PDF Layout with Banners)
# -------------------------------------------------------------
ws1 = wb.active
ws1.title = "मतदान केन्द्रों की सूची"
ws1.views.sheetView[0].showGridLines = True

# Title Block
ws1.merge_cells("A1:J1")
ws1["A1"] = "कोशी स्नातक निर्वाचन क्षेत्र के मतदान केन्द्रों की सूची"
ws1["A1"].font = font_title
ws1["A1"].fill = fill_title
ws1["A1"].alignment = align_center_middle
ws1.row_dimensions[1].height = 36

ws1.merge_cells("A2:J2")
ws1["A2"] = "बिहार विधान परिषद् निर्वाचन • कुल 14 जिले • 181 मतदान केन्द्र • 1,34,593 कुल मतदाता"
ws1["A2"].font = font_subtitle
ws1["A2"].fill = fill_subtitle
ws1["A2"].alignment = align_center_middle
ws1.row_dimensions[2].height = 24

ws1.row_dimensions[3].height = 8

# Table Headers (Row 4)
headers_sheet1 = [
    "क्र०",
    "मतदान केन्द्र का स्थान",
    "भवन जिसमें मतदान केन्द्र अवस्थित होगा",
    "मतदान केन्द्र का क्षेत्रफल",
    "क्या पृथक प्रवेश एवं निकास द्वार है। यदि नहीं तो तत्संबंधी कारण",
    "मतदान क्षेत्र की विवरणी",
    "क्या सभी मतदाताओं के लिए या केवल स्नातकों के लिए या शिक्षकों के लिए",
    "आवंटित मतदाताओं की कुल संख्या",
    "मतदान केन्द्र पहुँचने के लिए मतदाताओं द्वारा तय की जानेवाली अधिकतम दूरी",
    "अभ्युक्ति"
]

for col_idx, h_text in enumerate(headers_sheet1, 1):
    cell = ws1.cell(row=4, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center_middle
    cell.border = border_cell
ws1.row_dimensions[4].height = 42

current_r = 5
current_district = None

for rec in raw_records:
    # Check if district changed, insert District Banner Row
    if rec['district'] != current_district:
        current_district = rec['district']
        # Banner row
        ws1.merge_cells(start_row=current_r, start_column=1, end_row=current_r, end_column=10)
        banner_cell = ws1.cell(row=current_r, column=1, value=f"जिला - {current_district}")
        banner_cell.font = font_banner
        banner_cell.fill = fill_banner
        banner_cell.alignment = align_center_middle
        for c in range(1, 11):
            ws1.cell(row=current_r, column=c).border = border_banner
        ws1.row_dimensions[current_r].height = 28
        current_r += 1
    
    # Write data row
    row_vals = [
        rec['serial'],
        rec['location'],
        rec['building'],
        rec['area'],
        rec['gate'],
        rec['voter_area'],
        rec['voter_type'],
        rec['voters_assigned_str'],
        rec['max_distance'],
        rec['remarks']
    ]
    
    fill_to_use = fill_row_alt if (current_r % 2 == 0) else fill_row_white
    
    for c_idx, val in enumerate(row_vals, 1):
        cell = ws1.cell(row=current_r, column=c_idx, value=val)
        cell.font = font_data
        cell.fill = fill_to_use
        cell.border = border_cell
        # Alignments
        if c_idx in [1, 4, 5, 7, 8, 9]:
            cell.alignment = align_center_top
        else:
            cell.alignment = align_left_top
    
    ws1.row_dimensions[current_r].height = 54
    current_r += 1

# Total Row on Sheet 1
ws1.merge_cells(start_row=current_r, start_column=1, end_row=current_r, end_column=7)
tot_label = ws1.cell(row=current_r, column=1, value="कुल योग (181 मतदान केन्द्र • 14 जिले)")
tot_label.font = font_total
tot_label.fill = fill_total
tot_label.alignment = Alignment(horizontal="right", vertical="center")

tot_val = ws1.cell(row=current_r, column=8, value=134593)
tot_val.font = font_total
tot_val.fill = fill_total
tot_val.alignment = align_center_middle
tot_val.number_format = "#,##0"

for c in range(1, 11):
    ws1.cell(row=current_r, column=c).border = border_total
    if c not in [1, 8]:
        ws1.cell(row=current_r, column=c).fill = fill_total
ws1.row_dimensions[current_r].height = 30

# Column widths for Sheet 1
col_widths_s1 = [8, 30, 32, 18, 16, 38, 20, 20, 18, 36]
for idx, w in enumerate(col_widths_s1, 1):
    ws1.column_dimensions[get_column_letter(idx)].width = w


# -------------------------------------------------------------
# SHEET 2: डेटाबेस - जिलावार (Structured Tabular Master Data)
# -------------------------------------------------------------
ws2 = wb.create_sheet(title="डेटाबेस (जिलावार)")
ws2.views.sheetView[0].showGridLines = True

headers_sheet2 = [
    "क्र०",
    "जिला",
    "मतदान केन्द्र का स्थान",
    "भवन जिसमें मतदान केन्द्र अवस्थित होगा",
    "मतदान केन्द्र का क्षेत्रफल",
    "क्या पृथक प्रवेश एवं निकास द्वार है",
    "मतदान क्षेत्र की विवरणी",
    "मतदाता श्रेणी",
    "आवंटित मतदाताओं की संख्या (विवरण)",
    "कुल मतदाता (संख्या)",
    "मतदान केन्द्र पहुँचने की अधिकतम दूरी",
    "अभ्युक्ति"
]

for col_idx, h_text in enumerate(headers_sheet2, 1):
    cell = ws2.cell(row=1, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center_middle
    cell.border = border_cell
ws2.row_dimensions[1].height = 36

for idx, rec in enumerate(raw_records, 2):
    fill_to_use = fill_row_alt if (idx % 2 == 0) else fill_row_white
    row_vals = [
        rec['serial'],
        rec['district'],
        rec['location'],
        rec['building'],
        rec['area'],
        rec['gate'],
        rec['voter_area'],
        rec['voter_type'],
        rec['voters_assigned_str'],
        rec['voters_count'],
        rec['max_distance'],
        rec['remarks']
    ]
    for c_idx, val in enumerate(row_vals, 1):
        cell = ws2.cell(row=idx, column=c_idx, value=val)
        cell.font = font_data
        cell.fill = fill_to_use
        cell.border = border_cell
        if c_idx == 10:
            cell.alignment = align_right_top
            cell.number_format = "#,##0"
        elif c_idx in [1, 2, 5, 6, 8, 11]:
            cell.alignment = align_center_top
        else:
            cell.alignment = align_left_top
    ws2.row_dimensions[idx].height = 46

# Total Row for Sheet 2
tot_r2 = len(raw_records) + 2
ws2.merge_cells(start_row=tot_r2, start_column=1, end_row=tot_r2, end_column=9)
tot_c2 = ws2.cell(row=tot_r2, column=1, value="कुल योग (Total)")
tot_c2.font = font_total
tot_c2.fill = fill_total
tot_c2.alignment = Alignment(horizontal="right", vertical="center")

tot_num = ws2.cell(row=tot_r2, column=10, value=f"=SUM(J2:J{tot_r2-1})")
tot_num.font = font_total
tot_num.fill = fill_total
tot_num.alignment = Alignment(horizontal="right", vertical="center")
tot_num.number_format = "#,##0"

for c in range(1, 13):
    ws2.cell(row=tot_r2, column=c).border = border_total
    if c not in [1, 10]:
        ws2.cell(row=tot_r2, column=c).fill = fill_total
ws2.row_dimensions[tot_r2].height = 28

# Set filters & freeze panes for Sheet 2
ws2.auto_filter.ref = f"A1:L{tot_r2-1}"
ws2.freeze_panes = "C2"

col_widths_s2 = [8, 14, 30, 32, 18, 16, 38, 20, 20, 16, 18, 36]
for idx, w in enumerate(col_widths_s2, 1):
    ws2.column_dimensions[get_column_letter(idx)].width = w


# -------------------------------------------------------------
# SHEET 3: जिलावार सारांश (District-wise Summary Table)
# -------------------------------------------------------------
ws3 = wb.create_sheet(title="जिलावार सारांश")
ws3.views.sheetView[0].showGridLines = True

# Title Block
ws3.merge_cells("A1:G1")
ws3["A1"] = "कोशी स्नातक निर्वाचन क्षेत्र - जिलावार मतदान केन्द्र एवं मतदाता सारांश"
ws3["A1"].font = font_title
ws3["A1"].fill = fill_title
ws3["A1"].alignment = align_center_middle
ws3.row_dimensions[1].height = 36

headers_sheet3 = [
    "क्र०",
    "जिला का नाम",
    "मतदान केन्द्रों की संख्या",
    "प्रारंभिक म०के० सं०",
    "अंतिम म०के० सं०",
    "कुल मतदाता",
    "% मतदाता हिस्सेदारी"
]

for col_idx, h_text in enumerate(headers_sheet3, 1):
    cell = ws3.cell(row=3, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center_middle
    cell.border = border_cell
ws3.row_dimensions[3].height = 30

# Calculate district stats
dist_groups = defaultdict(list)
for rec in raw_records:
    dist_groups[rec['district']].append(rec)

cur_s3_row = 4
for idx, d_name in enumerate(canonical_districts, 1):
    recs = dist_groups[d_name]
    start_s = recs[0]['serial']
    end_s = recs[-1]['serial']
    tot_st = len(recs)
    tot_vt = sum(r['voters_count'] for r in recs)
    
    fill_to_use = fill_row_alt if (cur_s3_row % 2 == 0) else fill_row_white
    
    vals = [
        idx,
        d_name,
        tot_st,
        start_s,
        end_s,
        tot_vt,
        f"=F{cur_s3_row}/$F$18"
    ]
    
    for c_idx, val in enumerate(vals, 1):
        cell = ws3.cell(row=cur_s3_row, column=c_idx, value=val)
        cell.font = font_data_bold if c_idx in [2, 3, 6] else font_data
        cell.fill = fill_to_use
        cell.border = border_cell
        if c_idx == 6:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0"
        elif c_idx == 7:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.00%"
        else:
            cell.alignment = align_center_middle
            
    ws3.row_dimensions[cur_s3_row].height = 24
    cur_s3_row += 1

# Total Row for Summary
ws3.merge_cells(start_row=18, start_column=1, end_row=18, end_column=2)
tot_c3 = ws3.cell(row=18, column=1, value="कुल योग (Total)")
tot_c3.font = font_total
tot_c3.fill = fill_total
tot_c3.alignment = Alignment(horizontal="center", vertical="center")

c_st_tot = ws3.cell(row=18, column=3, value="=SUM(C4:C17)")
c_st_tot.font = font_total
c_st_tot.fill = fill_total
c_st_tot.alignment = align_center_middle

ws3.cell(row=18, column=4).fill = fill_total
ws3.cell(row=18, column=5).fill = fill_total

c_vt_tot = ws3.cell(row=18, column=6, value="=SUM(F4:F17)")
c_vt_tot.font = font_total
c_vt_tot.fill = fill_total
c_vt_tot.alignment = Alignment(horizontal="right", vertical="center")
c_vt_tot.number_format = "#,##0"

c_pct_tot = ws3.cell(row=18, column=7, value="=SUM(G4:G17)")
c_pct_tot.font = font_total
c_pct_tot.fill = fill_total
c_pct_tot.alignment = Alignment(horizontal="right", vertical="center")
c_pct_tot.number_format = "0.00%"

for c in range(1, 8):
    ws3.cell(row=18, column=c).border = border_total
ws3.row_dimensions[18].height = 28

col_widths_s3 = [8, 16, 22, 18, 18, 18, 18]
for idx, w in enumerate(col_widths_s3, 1):
    ws3.column_dimensions[get_column_letter(idx)].width = w

# Save Workbook
output_path = "Matdan_Kendron_Ki_Suchi.xlsx"
wb.save(output_path)
print(f"Successfully generated: {output_path}")
