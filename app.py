from flask import Flask, render_template, request, jsonify, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import barcode
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
import os
import io
import tempfile
from datetime import datetime

app = Flask(__name__)

# ═══════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════
JSON_FILE = "cloth-barcode-system-893c74d7606d.json"
SHEET_NAME = "Mission Factory"
OUTPUT_FOLDER = "generated_slips"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ═══════════════════════════════════════════════════════════════════
# PIECE LIST (same as desktop app)
# ═══════════════════════════════════════════════════════════════════
PIECE_LIST = {
    "0": ["000", "TEST / Missing"],
    "1": ["864", "T/N Plain"], "2": ["873", "T/N Selection"],
    "3": ["801", "RN Coty"], "4": ["702", "V/S Coty"],
    "5": ["9333", "V/S Pullover V-Mark"], "6": ["9360", "V/S S/L"],
    "7": ["9207", "Interatia T-Shirt"], "8": ["882", "Plain T-Shirt"],
    "9": ["8370", "12-GG T-Shirt"], "10": ["9306", "T-Shirt 1x1 Designer"],
    "11": ["1008", "T-Shirt Patta"], "12": ["927", "T/N Patta"],
    "13": ["1071", "T/N Zipper"], "14": ["900", "H/N Zipper"],
    "15": ["865", "H/N Plain"], "16": ["874", "H/N Selection"],
    "17": ["855", "V/S Pullover 12-GG"], "18": ["856", "V/S S/L 12-GG"],
    "19": ["8055", "V/S S/L Coat"], "20": ["4005", "V/S S/L Coty"],
    "21": ["603", "Allover V/S Pullover Coty"], "22": ["792", "V/S Coty V-Mark"],
    "23": ["7002", "V/S Coty Computer"], "24": ["6003", "V/S S/L Coty 42 Size"],
    "25": ["5001", "T/N Skivi 14/16/18"], "26": ["5010", "T/N Skivi 20/30"],
    "27": ["5031", "T/N Skivi 32/36"], "28": ["5100", "T/N Skivi Free Size"],
    "29": ["5013", "H/N Skivi 20/30"], "30": ["5032", "H/N Skivi 32/36"],
    "31": ["5101", "H/N Skivi Free Size"], "32": ["5058", "T/N Shimmer Skivi 20/30"],
    "33": ["5059", "T/N Shimmer Skivi 32/36"], "34": ["5060", "T/N Shimmer Skivi Free Size"],
    "35": ["501", "Pajami 18-2/20-2/22-2"], "36": ["502", "Pajami 24-2/26-2/28-1/30-1"],
    "37": ["503", "Pajami 32/36"], "38": ["", "Jersey Pullover"]
}

SKIVI_REFS = {str(i) for i in range(25, 38)} | {"00"}
TN_REFS = {str(i) for i in range(1, 25)} | {"0"} | {"38"}

# ═══════════════════════════════════════════════════════════════════
# GOOGLE SHEETS CONNECTION
# ═══════════════════════════════════════════════════════════════════
def connect_to_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds",
                 "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
        client = gspread.authorize(creds)
        ss = client.open(SHEET_NAME)
        tn_sheet = ss.worksheet("T/N Process")
        skivi_sheet = ss.worksheet("Skivi Process")
        form = ss.worksheet("Form Details")
        return tn_sheet, skivi_sheet, form, None
    except Exception as e:
        return None, None, None, str(e)

tn_process_sheet, skivi_process_sheet, form_sheet, _conn_error = connect_to_sheets()

# ═══════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════
def get_expected_slip_type(ref_no: str) -> str:
    return "Skivi" if ref_no in SKIVI_REFS else "T/N"

def slip_prefix(slip_type: str) -> str:
    return "S" if slip_type == "Skivi" else "T"

def check_duplicate_slip(slip_type: str, slip_no: str) -> bool:
    if not form_sheet:
        return False
    try:
        all_rows = form_sheet.get_all_values()
        prefix = slip_prefix(slip_type)
        for row in all_rows[1:]:
            if len(row) >= 3:
                existing_prefix = slip_prefix(row[2].strip())
                existing_no = row[1].strip()
                if existing_prefix == prefix and existing_no == slip_no:
                    return True
    except:
        pass
    return False

def get_first_empty_row(sheet):
    try:
        return len(sheet.col_values(1)) + 1
    except:
        return 2

def generate_barcode_tempfile(data: str) -> str:
    buf = io.BytesIO()
    barcode.generate("code128", data, writer=ImageWriter(), output=buf,
                     writer_options={"write_text": False, "quiet_zone": 2,
                                     "module_height": 10.0, "module_width": 0.2})
    buf.seek(0)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.write(buf.read())
    tmp.close()
    return tmp.name

# ═══════════════════════════════════════════════════════════════════
# PDF LAYOUT CONSTANTS (same as desktop)
# ═══════════════════════════════════════════════════════════════════
PAGE_W = 52 * mm
PAGE_H = 297 * mm
MARGIN = 2 * mm
CONTENT_W = PAGE_W - 2 * MARGIN
BLOCK_H_TN = 15.38 * mm
BLOCK_H_SKIVI = 23.75 * mm
_INFO_H_MM = 3.0
_ID_H_MM = 2.5
_PROC_H_MM = 4.0

def _block_dims(slip_type: str) -> tuple:
    block_h = BLOCK_H_TN if slip_type == "T/N" else BLOCK_H_SKIVI
    info_h = _INFO_H_MM * mm
    id_h = _ID_H_MM * mm
    proc_h = _PROC_H_MM * mm
    bar_h = block_h - info_h - id_h - proc_h
    return block_h, info_h, bar_h, id_h, proc_h

# ═══════════════════════════════════════════════════════════════════
# PDF BUILDER (same logic as desktop)
# ═══════════════════════════════════════════════════════════════════
def build_slip_pdf(slip_type, slip_no, ref_no, pcs, p_name, all_processes, now) -> str:
    safe_type = slip_type.replace("/", "-")
    filename = f"{safe_type}-{slip_no}_{now.strftime('%d-%m-%Y_%H-%M-%S')}.pdf"
    filepath = os.path.join(OUTPUT_FOLDER, filename)

    block_h, info_h, bar_h, id_h, proc_h = _block_dims(slip_type)
    n_blocks = len(all_processes)
    page_h = MARGIN + block_h * n_blocks + MARGIN

    c = rl_canvas.Canvas(filepath, pagesize=(PAGE_W, page_h))
    tmpfiles = []
    y = page_h - MARGIN

    for proc_name, is_active in all_processes:
        info_top = y
        info_bot = info_top - info_h
        bar_top = info_bot
        bar_bot = bar_top - bar_h
        id_top = bar_bot
        id_bot = id_top - id_h
        proc_top = id_bot
        proc_bot = proc_top - proc_h

        info_baseline = info_bot + info_h * 0.35
        c.setFont("Helvetica-Bold", 5)
        c.setFillColor(colors.black)
        c.drawString(MARGIN, info_baseline, f"Ref: {ref_no} | Pcs: {pcs} | {p_name}")
        c.setFont("Helvetica-Bold", 7)
        c.drawRightString(PAGE_W - MARGIN, info_baseline, f"*{slip_no}*")

        if is_active:
            unique_id = f"{slip_prefix(slip_type)}{slip_no}R{ref_no}Q{pcs}"
            try:
                bar_path = generate_barcode_tempfile(unique_id)
                tmpfiles.append(bar_path)
                c.drawImage(bar_path, MARGIN, bar_bot, width=CONTENT_W, height=bar_h,
                            preserveAspectRatio=False)
            except Exception as e:
                c.setFillColor(colors.red)
                c.setFont("Helvetica", 5)
                c.drawString(MARGIN, bar_bot + bar_h * 0.4, f"Barcode error: {e}")
                c.setFillColor(colors.black)

            id_baseline = id_bot + id_h * 0.35
            c.setFont("Helvetica", 4.5)
            c.setFillColor(colors.black)
            c.drawCentredString(PAGE_W / 2, id_baseline, unique_id)
        else:
            c.setStrokeColor(colors.HexColor("#DDDDDD"))
            c.setFillColor(colors.HexColor("#F8F8F8"))
            c.setLineWidth(0.3)
            c.rect(MARGIN, bar_bot, CONTENT_W, bar_h, stroke=1, fill=1)
            c.setFillColor(colors.HexColor("#AAAAAA"))
            c.setFont("Helvetica", 5)
            c.drawCentredString(PAGE_W / 2, bar_bot + bar_h * 0.4, "INACTIVE")
            c.setFillColor(colors.black)

        proc_baseline = proc_bot + proc_h * 0.30
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.black)
        c.drawCentredString(PAGE_W / 2, proc_baseline, f"--- {proc_name.upper()} ---")
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.8)
        c.line(MARGIN, proc_bot, PAGE_W - MARGIN, proc_bot)
        y = proc_bot

    c.save()
    for tf in tmpfiles:
        try:
            os.remove(tf)
        except:
            pass
    return filepath

# ═══════════════════════════════════════════════════════════════════
# WEB ROUTES
# ═══════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html', piece_list=PIECE_LIST)

@app.route('/get_piece_info/<ref_no>')
def get_piece_info(ref_no):
    if ref_no not in PIECE_LIST:
        return jsonify({"error": "Invalid reference number"}), 400
    
    info = PIECE_LIST[ref_no]
    expected_type = get_expected_slip_type(ref_no)
    
    return jsonify({
        "code": info[0],
        "name": info[1],
        "expected_type": expected_type
    })

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    slip_type = data.get('slip_type')
    slip_no = data.get('slip_no')
    ref_no = data.get('ref_no')
    pcs = data.get('pcs')

    # Validation
    if not all([slip_type, slip_no, ref_no, pcs]):
        return jsonify({"error": "All fields are required"}), 400

    if ref_no not in PIECE_LIST:
        return jsonify({"error": "Invalid piece reference number"}), 400

    expected = get_expected_slip_type(ref_no)
    if slip_type != expected:
        p_name = f"{PIECE_LIST[ref_no][0]} {PIECE_LIST[ref_no][1]}"
        return jsonify({
            "error": f"Ref {ref_no} ({p_name}) belongs to '{expected}' slips only. You selected '{slip_type}'."
        }), 400

    if check_duplicate_slip(slip_type, slip_no):
        prefix = slip_prefix(slip_type)
        return jsonify({
            "error": f"Slip {prefix}-{slip_no} has already been used! Please use a different slip number."
        }), 400

    # Generate PDF
    try:
        now = datetime.now()
        p_info = PIECE_LIST[ref_no]
        p_name = f"{p_info[0]} {p_info[1]}"

        # Log to Google Sheets
        if form_sheet:
            row = get_first_empty_row(form_sheet)
            form_sheet.update(
                f"A{row}:F{row}",
                [[now.strftime("%Y-%m-%d %H:%M:%S"), slip_no, slip_type, ref_no, pcs, "Opened"]],
                value_input_option="USER_ENTERED"
            )

        # Get processes
        all_processes = []
        process_sheet = tn_process_sheet if slip_type == "T/N" else skivi_process_sheet
        if process_sheet:
            cell = process_sheet.find(ref_no)
            p_row = process_sheet.row_values(cell.row)
            headers = process_sheet.row_values(1)
            for i in range(2, len(headers)):
                proc_name = headers[i]
                if not proc_name.strip():
                    continue
                cell_val = p_row[i].strip().upper() if i < len(p_row) else "FALSE"
                is_active = (cell_val == "TRUE")
                all_processes.append((proc_name, is_active))

        if not all_processes:
            return jsonify({"error": f"No processes found for Ref {ref_no}"}), 400

        # Build PDF
        pdf_path = build_slip_pdf(slip_type, slip_no, ref_no, pcs, p_name, all_processes, now)
        
        return jsonify({
            "success": True,
            "message": "PDF generated successfully!",
            "filename": os.path.basename(pdf_path)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({"error": "File not found"}), 404

@app.route('/health')
def health():
    sheets_connected = form_sheet is not None
    return jsonify({
        "status": "ok",
        "sheets_connected": sheets_connected
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)