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

def get_block(serial, dist):
    # Khagaria (1 - 8)
    if dist == 'खगड़िया':
        if serial in ['1', '1क']: return 'अलौली'
        if serial in ['2', '2क']: return 'खगड़िया (ग्रामीण)'
        if serial == '3': return 'खगड़िया (नगर परिषद्)'
        if serial == '4': return 'मानसी'
        if serial == '5': return 'चैथम'
        if serial == '6': return 'बेलदौर'
        if serial in ['7', '7क']: return 'गोगरी'
        if serial == '8': return 'परबत्ता'

    # Saharsa (9 - 19)
    if dist == 'सहरसा':
        if serial == '9': return 'नवहट्टा'
        if serial == '10': return 'महिषी'
        if serial in ['11', '11(क)']: return 'सहरसा (नगर निगम)'
        if serial == '12': return 'कहरा'
        if serial == '13': return 'सत्तरकटैया'
        if serial == '14': return 'सौरबाजार'
        if serial == '15': return 'पतरघट'
        if serial == '16': return 'सोनवर्षा'
        if serial in ['17', '17(क)']: return 'सिमरी बख्तियारपुर'
        if serial == '18': return 'सलखुआ'
        if serial == '19': return 'बनमा ईटहरी'

    # Madhepura (20 - 33)
    if dist == 'मधेपुरा':
        if serial == '20': return 'घैलाढ़'
        if serial == '21': return 'गम्हरिया'
        if serial == '22': return 'सिंहेश्वर'
        if serial in ['23', '23क']: return 'मधेपुरा (सदर)'
        if serial == '24': return 'शंकरपुर'
        if serial in ['25', '25क']: return 'मुरलीगंज'
        if serial == '26': return 'कुमारखंड'
        if serial == '27': return 'ग्वालपाड़ा'
        if serial == '28': return 'बिहारीगंज'
        if serial == '29': return 'किशुनगंज'
        if serial == '30': return 'उदाकिशुनगंज'
        if serial == '31': return 'आलमनगर'
        if serial == '32': return 'पुरैनी'
        if serial == '33': return 'चौसा'

    # Supaul (34 - 45)
    if dist == 'सुपौल':
        if serial == '34': return 'मरौना'
        if serial == '35': return 'निर्मली'
        if serial == '36': return 'सरायगढ़ भपटियाही'
        if serial == '37': return 'किशनपुर'
        if serial in ['38', '38(क)']: return 'सुपौल (सदर)'
        if serial == '39': return 'पिपरा'
        if serial == '40': return 'राघोपुर'
        if serial == '41': return 'प्रतापगंज'
        if serial == '42': return 'बसंतपुर'
        if serial == '43': return 'छातापुर (भाग-1)'
        if serial == '44': return 'छातापुर (भाग-2)'
        if serial in ['45', '45 (क)']: return 'त्रिवेणीगंज'

    # Araria (46 - 56)
    if dist == 'अररिया':
        if serial == '46': return 'नरपतगंज'
        if serial == '47': return 'फारबिसगंज (ग्रामीण)'
        if serial == '48': return 'फारबिसगंज (नगर परिषद्)'
        if serial == '49': return 'भरगामा'
        if serial in ['50', '50क']: return 'रानीगंज'
        if serial == '51': return 'अररिया (सदर)'
        if serial == '52': return 'अररिया (नगर परिषद्)'
        if serial in ['53', '53(क)']: return 'जोकीहाट'
        if serial == '54': return 'कुर्साकांटा'
        if serial == '55': return 'सिकटी'
        if serial == '56': return 'पलासी'

    # Kishanganj (57 - 64)
    if dist == 'किशनगंज':
        if serial == '57': return 'बहादुरगंज'
        if serial == '58': return 'दिघलबैंक'
        if serial == '59': return 'ठाकुरगंज'
        if serial == '60': return 'पोठिया'
        if serial == '61': return 'कोचाधामन'
        if serial == '62': return 'टेढ़ागाछ'
        if serial == '63': return 'किशनगंज (ग्रामीण)'
        if serial == '64': return 'किशनगंज (नगर परिषद्)'

    # Purnia (65 - 79)
    if dist == 'पूर्णिया':
        if serial == '65': return 'बनमनखी'
        if serial == '66': return 'बीकोठी'
        if serial == '67': return 'भवानीपुर'
        if serial == '68': return 'रूपौली'
        if serial == '69': return 'धमदाहा'
        if serial == '70': return 'कृत्यानंदनगर'
        if serial == '71': return 'पूर्णिया पूर्व'
        if serial in ['72', '72 क', '72 ख']: return 'पूर्णिया (नगर निगम)'
        if serial == '73': return 'कसबा'
        if serial in ['74', '74क']: return 'जलालगढ़'
        if serial == '75': return 'श्रीनगर'
        if serial == '76': return 'डगरुआ'
        if serial in ['77', '77क']: return 'बायसी'
        if serial == '78': return 'अमौर'
        if serial == '79': return 'बैसा'

    # Katihar (80 - 97)
    if dist == 'कटिहार':
        if serial == '80': return 'फलका'
        if serial == '81': return 'समेली'
        if serial == '82': return 'कुर्सेला'
        if serial in ['83', '83 (क)']: return 'बरारी'
        if serial == '84': return 'कोढ़ा'
        if serial in ['85', '85 (क)']: return 'कटिहार (सदर)'
        if serial in ['86', '86 (क)']: return 'कटिहार (नगर निगम)'
        if serial == '87': return 'हसनगंज'
        if serial == '88': return 'डंडखोरा'
        if serial in ['89', '89 (क)']: return 'मनसाही'
        if serial == '90': return 'प्राणपुर'
        if serial == '91': return 'बलरामपुर'
        if serial == '92': return 'बारसोई'
        if serial == '93': return 'कदवा'
        if serial == '94': return 'आजमनगर'
        if serial == '95': return 'मनिहारी'
        if serial == '96': return 'अमदाबाद'
        if serial == '97': return 'बारसोई'

    # Bhagalpur (98 - 116)
    if dist == 'भागलपुर':
        if serial == '98': return 'गोराडीह'
        if serial == '99': return 'नारायणपुर'
        if serial == '100': return 'बिहपुर'
        if serial == '101': return 'सुलतानगंज'
        if serial == '102': return 'शाहकुण्ड'
        if serial == '103': return 'नाथनगर'
        if serial == '104': return 'जगदीशपुर'
        if serial in ['105', '105-क']: return 'सबौर'
        if serial in ['106', '106-क', '107', '107-क', '107-ख', '108', '108-क']: return 'भागलपुर (नगर निगम)'
        if serial == '109': return 'खरीक'
        if serial in ['110', '111']: return 'नवगछिया'
        if serial == '112': return 'गोपालपुर'
        if serial == '113': return 'रंगरा चौक'
        if serial in ['114', '114-क']: return 'पीरपैंती'
        if serial in ['115', '115-क']: return 'कहलगाँव'
        if serial == '116': return 'सन्हौला'

    # Munger (117 - 127)
    if dist == 'मुंगेर':
        if serial == '117': return 'मुंगेर (सदर)'
        if serial in ['118', '118क']: return 'मुंगेर (नगर निगम)'
        if serial == '119': return 'बरियारपुर'
        if serial == '120': return 'जमालपुर (नगर परिषद्)'
        if serial == '121': return 'जमालपुर (ग्रामीण)'
        if serial == '122': return 'धरहरा'
        if serial == '123': return 'हवेली खड़गपुर'
        if serial == '124': return 'असरगंज'
        if serial == '125': return 'तारापुर'
        if serial == '126': return 'टेटिया बम्बर'
        if serial == '127': return 'संग्रामपुर'

    # Lakhisarai (128 - 134)
    if dist == 'लखीसराय':
        if serial == '128': return 'बड़हिया'
        if serial == '129': return 'पिपरिया'
        if serial == '130': return 'सूर्यगढ़ा'
        if serial == '131': return 'चानन'
        if serial == '132': return 'लखीसराय (नगर परिषद)'
        if serial == '133': return 'लखीसराय (ग्रामीण)'
        if serial == '134': return 'हलसी एवं रामगढ़चौक'

    # Sheikhpura (135 - 141)
    if dist == 'शेखपुरा':
        if serial == '135': return 'बरबीघा'
        if serial == '136': return 'शेखोपुरसराय'
        if serial == '137': return 'शेखपुरा (ग्रामीण)'
        if serial == '138': return 'घाटकुसुम्भा'
        if serial == '139': return 'शेखपुरा (नगर परिषद)'
        if serial == '140': return 'अरियरी'
        if serial == '141': return 'चेवाड़ा'

    # Jamui (142 - 152)
    if dist == 'जमुई':
        if serial == '142': return 'ई० अलीगंज'
        if serial == '143': return 'सिकन्दरा'
        if serial == '144': return 'खैरा'
        if serial == '145': return 'जमुई (ग्रामीण)'
        if serial == '146': return 'जमुई (नगर परिषद्)'
        if serial == '147': return 'बरहट'
        if serial == '148': return 'लक्ष्मीपुर'
        if serial == '149': return 'गिद्धौर'
        if serial == '150': return 'झाझा'
        if serial == '151': return 'सोनो'
        if serial == '152': return 'चकाई'

    # Banka (153 - 163)
    if dist == 'बाँका':
        if serial == '153': return 'शंभूगंज'
        if serial == '154': return 'अमरपुर'
        if serial == '155': return 'रजौन'
        if serial == '156': return 'धोरैया'
        if serial == '157': return 'बाराहाट'
        if serial == '158': return 'बाँका (सदर)'
        if serial == '159': return 'फुल्लीडुमर'
        if serial == '160': return 'बेलहर'
        if serial == '161': return 'चांदन'
        if serial == '162': return 'कटोरिया'
        if serial == '163': return 'बौंसी'

    return 'अन्य'

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
                while len(cleaned_cells) < 10:
                    cleaned_cells.append("")
                row_full_text = " ".join(cleaned_cells)
                dist = detect_district(row_full_text)
                blk = get_block(cleaned_cells[0], dist)
                raw_records.append({
                    'page': page_idx + 1,
                    'row_idx': row_idx,
                    'serial': cleaned_cells[0],
                    'district': dist,
                    'block': blk,
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
# SHEET 1: मतदान केन्द्रों की सूची (Exact PDF Layout with Block & Banners)
# -------------------------------------------------------------
ws1 = wb.active
ws1.title = "मतदान केन्द्रों की सूची"
ws1.views.sheetView[0].showGridLines = True

# Title Block
ws1.merge_cells("A1:K1")
ws1["A1"] = "कोशी स्नातक निर्वाचन क्षेत्र के मतदान केन्द्रों की सूची"
ws1["A1"].font = font_title
ws1["A1"].fill = fill_title
ws1["A1"].alignment = align_center_middle
ws1.row_dimensions[1].height = 36

ws1.merge_cells("A2:K2")
ws1["A2"] = "बिहार विधान परिषद् निर्वाचन • कुल 14 जिले • 181 मतदान केन्द्र • 1,34,593 कुल मतदाता"
ws1["A2"].font = font_subtitle
ws1["A2"].fill = fill_subtitle
ws1["A2"].alignment = align_center_middle
ws1.row_dimensions[2].height = 24

ws1.row_dimensions[3].height = 8

# Table Headers (Row 4) - 11 columns with Block
headers_sheet1 = [
    "क्र०",
    "प्रखंड / क्षेत्र",
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
        # Banner row across all 11 columns
        ws1.merge_cells(start_row=current_r, start_column=1, end_row=current_r, end_column=11)
        banner_cell = ws1.cell(row=current_r, column=1, value=f"जिला - {current_district}")
        banner_cell.font = font_banner
        banner_cell.fill = fill_banner
        banner_cell.alignment = align_center_middle
        for c in range(1, 12):
            ws1.cell(row=current_r, column=c).border = border_banner
        ws1.row_dimensions[current_r].height = 28
        current_r += 1
    
    # Write data row
    row_vals = [
        rec['serial'],
        rec['block'],
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
        if c_idx in [1, 2, 5, 6, 8, 9, 10]:
            cell.alignment = align_center_top
        else:
            cell.alignment = align_left_top
    
    ws1.row_dimensions[current_r].height = 54
    current_r += 1

# Total Row on Sheet 1
ws1.merge_cells(start_row=current_r, start_column=1, end_row=current_r, end_column=8)
tot_label = ws1.cell(row=current_r, column=1, value="कुल योग (181 मतदान केन्द्र • 14 जिले)")
tot_label.font = font_total
tot_label.fill = fill_total
tot_label.alignment = Alignment(horizontal="right", vertical="center")

tot_val = ws1.cell(row=current_r, column=9, value=134593)
tot_val.font = font_total
tot_val.fill = fill_total
tot_val.alignment = align_center_middle
tot_val.number_format = "#,##0"

for c in range(1, 12):
    ws1.cell(row=current_r, column=c).border = border_total
    if c not in [1, 9]:
        ws1.cell(row=current_r, column=c).fill = fill_total
ws1.row_dimensions[current_r].height = 30

# Column widths for Sheet 1 (11 columns)
col_widths_s1 = [8, 18, 30, 32, 18, 16, 38, 20, 20, 18, 36]
for idx, w in enumerate(col_widths_s1, 1):
    ws1.column_dimensions[get_column_letter(idx)].width = w


# -------------------------------------------------------------
# SHEET 2: डेटाबेस - जिलावार एवं प्रखंडवार (Master Data)
# -------------------------------------------------------------
ws2 = wb.create_sheet(title="डेटाबेस (जिलावार)")
ws2.views.sheetView[0].showGridLines = True

headers_sheet2 = [
    "क्र०",
    "जिला",
    "प्रखंड / क्षेत्र",
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
        rec['block'],
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
        if c_idx == 11:
            cell.alignment = align_right_top
            cell.number_format = "#,##0"
        elif c_idx in [1, 2, 3, 6, 7, 9, 12]:
            cell.alignment = align_center_top
        else:
            cell.alignment = align_left_top
    ws2.row_dimensions[idx].height = 46

# Total Row for Sheet 2
tot_r2 = len(raw_records) + 2
ws2.merge_cells(start_row=tot_r2, start_column=1, end_row=tot_r2, end_column=10)
tot_c2 = ws2.cell(row=tot_r2, column=1, value="कुल योग (Total)")
tot_c2.font = font_total
tot_c2.fill = fill_total
tot_c2.alignment = Alignment(horizontal="right", vertical="center")

tot_num = ws2.cell(row=tot_r2, column=11, value=f"=SUM(K2:K{tot_r2-1})")
tot_num.font = font_total
tot_num.fill = fill_total
tot_num.alignment = Alignment(horizontal="right", vertical="center")
tot_num.number_format = "#,##0"

for c in range(1, 14):
    ws2.cell(row=tot_r2, column=c).border = border_total
    if c not in [1, 11]:
        ws2.cell(row=tot_r2, column=c).fill = fill_total
ws2.row_dimensions[tot_r2].height = 28

# Set filters & freeze panes for Sheet 2
ws2.auto_filter.ref = f"A1:M{tot_r2-1}"
ws2.freeze_panes = "D2"

col_widths_s2 = [8, 14, 18, 30, 32, 18, 16, 38, 20, 20, 16, 18, 36]
for idx, w in enumerate(col_widths_s2, 1):
    ws2.column_dimensions[get_column_letter(idx)].width = w


# -------------------------------------------------------------
# SHEET 3: जिलावार एवं प्रखंडवार सारांश (District & Block Summary)
# -------------------------------------------------------------
ws3 = wb.create_sheet(title="जिलावार सारांश")
ws3.views.sheetView[0].showGridLines = True

# Title Block
ws3.merge_cells("A1:H1")
ws3["A1"] = "कोशी स्नातक निर्वाचन क्षेत्र - जिलावार व प्रखंडवार मतदान केन्द्र एवं मतदाता सारांश"
ws3["A1"].font = font_title
ws3["A1"].fill = fill_title
ws3["A1"].alignment = align_center_middle
ws3.row_dimensions[1].height = 36

headers_sheet3 = [
    "क्र०",
    "जिला का नाम",
    "प्रखंडों / क्षेत्रों की संख्या",
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
    unique_blks = len(set(r['block'] for r in recs))
    
    fill_to_use = fill_row_alt if (cur_s3_row % 2 == 0) else fill_row_white
    
    vals = [
        idx,
        d_name,
        unique_blks,
        tot_st,
        start_s,
        end_s,
        tot_vt,
        f"=G{cur_s3_row}/$G$18"
    ]
    
    for c_idx, val in enumerate(vals, 1):
        cell = ws3.cell(row=cur_s3_row, column=c_idx, value=val)
        cell.font = font_data_bold if c_idx in [2, 3, 4, 7] else font_data
        cell.fill = fill_to_use
        cell.border = border_cell
        if c_idx == 7:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0"
        elif c_idx == 8:
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

c_blk_tot = ws3.cell(row=18, column=3, value="=SUM(C4:C17)")
c_blk_tot.font = font_total
c_blk_tot.fill = fill_total
c_blk_tot.alignment = align_center_middle

c_st_tot = ws3.cell(row=18, column=4, value="=SUM(D4:D17)")
c_st_tot.font = font_total
c_st_tot.fill = fill_total
c_st_tot.alignment = align_center_middle

ws3.cell(row=18, column=5).fill = fill_total
ws3.cell(row=18, column=6).fill = fill_total

c_vt_tot = ws3.cell(row=18, column=7, value="=SUM(G4:G17)")
c_vt_tot.font = font_total
c_vt_tot.fill = fill_total
c_vt_tot.alignment = Alignment(horizontal="right", vertical="center")
c_vt_tot.number_format = "#,##0"

c_pct_tot = ws3.cell(row=18, column=8, value="=SUM(H4:H17)")
c_pct_tot.font = font_total
c_pct_tot.fill = fill_total
c_pct_tot.alignment = Alignment(horizontal="right", vertical="center")
c_pct_tot.number_format = "0.00%"

for c in range(1, 9):
    ws3.cell(row=18, column=c).border = border_total
ws3.row_dimensions[18].height = 28

col_widths_s3 = [8, 16, 24, 22, 18, 18, 18, 18]
for idx, w in enumerate(col_widths_s3, 1):
    ws3.column_dimensions[get_column_letter(idx)].width = w

# Save Workbook
output_path = "Matdan_Kendron_Ki_Suchi.xlsx"
wb.save(output_path)
print(f"Successfully updated: {output_path} with Block column!")
