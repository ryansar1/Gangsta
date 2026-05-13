import streamlit as st
import streamlit.components.v1 as components
import csv, io, openpyxl
import qrcode
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="Loaner Labels", page_icon="🏷️", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    h1, h2, h3 { font-family: 'Barlow Condensed', sans-serif !important; letter-spacing: 0.5px; }
    .stButton > button {
        background-color: #1a3a5c; color: white;
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 18px; font-weight: 600; letter-spacing: 1px;
        border: none; border-radius: 4px; padding: 0.6em 2em; width: 100%;
    }
    .stButton > button:hover { background-color: #122840; color: white; }
    .download-bar { background:#eef4ff; border:1px solid #1a3a5c; border-radius:6px; padding:14px 18px; margin-bottom:16px; }
    .download-bar-title { font-family:'Barlow Condensed',sans-serif; font-size:15px; font-weight:600; color:#1a3a5c; margin-bottom:10px; }
    .stat-box { background:white; border-left:4px solid #1a3a5c; border-radius:4px; padding:14px 18px; margin:8px 0; box-shadow:0 1px 4px rgba(0,0,0,0.08); }
    .stat-number { font-family:'Barlow Condensed',sans-serif; font-size:34px; font-weight:700; color:#1a3a5c; line-height:1; }
    .stat-label  { font-size:13px; color:#666; margin-top:4px; text-transform:uppercase; letter-spacing:0.5px; }
    .upload-hint { text-align:center; color:#888; font-size:13px; margin-top:-8px; margin-bottom:16px; }
</style>
""", unsafe_allow_html=True)

st.markdown("# 🏷️ Loaner Label Generator")
st.markdown("Upload your vehicle list, pick your start slot, and download print-ready Avery 8460 labels.")
st.divider()

# ── Constants ─────────────────────────────────────────────────
COLS        = 3
ROWS        = 10
SLOTS       = COLS * ROWS
LABEL_W     = 2.625 * inch
LABEL_H     = 1.0   * inch
MARGIN_LEFT = 0.19  * inch
MARGIN_TOP  = 0.5   * inch
H_GAP       = 0.125 * inch
V_GAP       = 0.0   * inch

NAVY       = colors.HexColor('#1a3a5c')
LIGHT_GRAY = colors.HexColor('#efefef')
MID_GRAY   = colors.HexColor('#cccccc')
DARK_GRAY  = colors.HexColor('#f2f2f2')
DARK_TEXT  = colors.HexColor('#111111')
MUTED_TEXT = colors.HexColor('#777777')
RULE_COLOR = colors.HexColor('#dedede')

# ── Session state ─────────────────────────────────────────────
for k, v in [('lbl_vehicles',[]),('lbl_start',0),('lbl_pdf_ready',False),('lbl_pdf_bytes',None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Parse file ────────────────────────────────────────────────
def parse_file(uploaded):
    vehicles = []
    name     = uploaded.name
    content  = uploaded.read()
    if name.endswith('.csv'):
        reader = csv.DictReader(io.StringIO(content.decode('utf-8-sig')))
        for row in reader:
            plate = row.get('License Plate', '').strip()
            vehicles.append({
                'stock': str(row.get('Unit Number', '')).strip(),
                'year':  str(row.get('Year', '')).strip(),
                'make':  str(row.get('Make', '')).strip().title(),
                'model': str(row.get('Model', '')).strip().title(),
                'vin':   str(row.get('VIN', '')).strip(),
                'plate': f'AZ · {plate}' if plate else '',
            })
    elif name.endswith(('.xlsx', '.xls')):
        wb      = openpyxl.load_workbook(io.BytesIO(content))
        ws      = wb.active
        headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
        for row in ws.iter_rows(min_row=2, values_only=True):
            r     = dict(zip(headers, [str(v).strip() if v is not None else '' for v in row]))
            plate = r.get('License Plate', '').strip()
            vehicles.append({
                'stock': r.get('Unit Number', ''),
                'year':  r.get('Year', ''),
                'make':  r.get('Make', '').title(),
                'model': r.get('Model', '').title(),
                'vin':   r.get('VIN', ''),
                'plate': f'AZ · {plate}' if plate else '',
            })
    return vehicles

# ── QR code ───────────────────────────────────────────────────
def make_qr(data):
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=5, border=2)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
    img = img.resize((180, 180), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return ImageReader(buf)

# ── Draw label ────────────────────────────────────────────────
def draw_label(c, x, y, v):
    W, H   = LABEL_W, LABEL_H
    pad    = 0.055 * inch
    qr_col = 0.88  * inch
    txt_w  = W - qr_col

    c.setStrokeColor(colors.HexColor('#222222'))
    c.setLineWidth(0.6)
    c.rect(x, y, W, H)

    hdr_h = 0.112 * inch
    hdr_y = y + H - hdr_h
    c.setFillColor(NAVY)
    c.rect(x, hdr_y, txt_w, hdr_h, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont('Helvetica-Bold', 5.5)
    c.drawString(x + pad, hdr_y + 0.031 * inch, 'LOANER VEHICLE')

    stk_h = 0.252 * inch
    stk_y = hdr_y - stk_h
    lbl_w = 0.30  * inch
    num_x = x + lbl_w + 0.01 * inch

    c.setFillColor(LIGHT_GRAY)
    c.rect(x, stk_y, txt_w, stk_h, stroke=0, fill=1)
    c.setStrokeColor(MID_GRAY); c.setLineWidth(0.5)
    c.line(x, stk_y, x + txt_w, stk_y)
    c.setStrokeColor(MID_GRAY); c.setLineWidth(0.4)
    c.line(x + lbl_w, stk_y, x + lbl_w, stk_y + stk_h)

    c.setFillColor(MUTED_TEXT)
    c.setFont('Helvetica-Bold', 6)
    c.drawCentredString(x + lbl_w / 2, stk_y + (stk_h / 2) + 0.012 * inch, 'STOCK')
    c.setFont('Helvetica', 5)
    c.drawCentredString(x + lbl_w / 2, stk_y + (stk_h / 2) - 0.022 * inch, '#')

    c.setFillColor(NAVY)
    c.setFont('Helvetica-Bold', 18)
    c.drawString(num_x, stk_y + 0.044 * inch, str(v.get('stock', '')))

    margin_bot = 0.016 * inch
    fields_top = stk_y
    fields_bot = y + margin_bot
    field_area = fields_top - fields_bot
    row_h      = field_area / 3

    def draw_field(label, value, fx, fy, fs=8.0):
        c.setFillColor(MUTED_TEXT)
        c.setFont('Helvetica', 4.5)
        c.drawString(fx, fy + row_h - 0.040 * inch, label)
        c.setFillColor(DARK_TEXT)
        c.setFont('Helvetica-Bold', fs)
        c.drawString(fx, fy + 0.014 * inch, str(value))

    r1_y = fields_bot + row_h * 2
    draw_field('YEAR',  v.get('year',  ''), x + pad,                r1_y, fs=8.0)
    draw_field('MAKE',  v.get('make',  ''), x + pad + 0.30 * inch,  r1_y, fs=8.0)
    draw_field('MODEL', v.get('model', ''), x + pad + 0.67 * inch,  r1_y, fs=7.0)
    c.setStrokeColor(RULE_COLOR); c.setLineWidth(0.3)
    c.line(x + pad, r1_y, x + txt_w - pad * 0.5, r1_y)

    r2_y = fields_bot + row_h
    draw_field('VIN', v.get('vin', ''), x + pad, r2_y, fs=6.0)
    c.setStrokeColor(RULE_COLOR); c.setLineWidth(0.3)
    c.line(x + pad, r2_y, x + txt_w - pad * 0.5, r2_y)

    r3_y = fields_bot
    draw_field('LICENSE PLATE', v.get('plate', ''), x + pad, r3_y, fs=8.5)

    c.setStrokeColor(MID_GRAY); c.setLineWidth(0.5)
    c.line(x + txt_w, y, x + txt_w, y + H)

    qr_x    = x + txt_w
    agr_h   = 0.30 * inch
    qr_zone = H - agr_h
    qr_pad  = 0.04 * inch
    qr_size = min(qr_col - qr_pad * 2, qr_zone - qr_pad * 2)
    qr_dx   = qr_x + (qr_col - qr_size) / 2
    qr_dy   = y + agr_h + (qr_zone - qr_size) / 2
    c.drawImage(make_qr(str(v.get('vin', 'NO-VIN'))), qr_dx, qr_dy, width=qr_size, height=qr_size)

    c.setFillColor(MUTED_TEXT); c.setFont('Helvetica', 4.0)
    c.drawCentredString(qr_x + qr_col / 2, y + agr_h + 0.022 * inch, 'SCAN FOR VIN')

    c.setStrokeColor(MID_GRAY); c.setLineWidth(0.5)
    c.line(qr_x, y + agr_h, qr_x + qr_col, y + agr_h)
    c.setFillColor(DARK_GRAY)
    c.rect(qr_x, y, qr_col, agr_h, stroke=0, fill=1)

    c.setFillColor(MUTED_TEXT); c.setFont('Helvetica-Bold', 4.5)
    c.drawCentredString(qr_x + qr_col / 2, y + agr_h - 0.060 * inch, 'AGREEMENT')

    blank_w = 0.15 * inch
    gap     = 0.05 * inch
    of_w    = 0.07 * inch
    total_w = blank_w + gap + of_w + gap + blank_w
    bx      = qr_x + (qr_col - total_w) / 2
    line_y  = y + 0.085 * inch
    c.setStrokeColor(DARK_TEXT); c.setLineWidth(1.0)
    c.line(bx, line_y, bx + blank_w, line_y)
    c.setFillColor(DARK_TEXT); c.setFont('Helvetica-Bold', 6.0)
    c.drawCentredString(bx + blank_w + gap + of_w / 2, line_y - 0.003 * inch, 'of')
    bx2 = bx + blank_w + gap + of_w + gap
    c.setStrokeColor(DARK_TEXT); c.setLineWidth(1.0)
    c.line(bx2, line_y, bx2 + blank_w, line_y)

# ── Generate PDF ──────────────────────────────────────────────
def generate_pdf(vehicles, start_slot):
    pw, ph    = letter
    buf       = io.BytesIO()
    c         = canvas.Canvas(buf, pagesize=letter)
    page_slot = start_slot
    for v in vehicles:
        if page_slot >= SLOTS:
            c.showPage()
            page_slot = 0
        col = page_slot % COLS
        row = page_slot // COLS
        lx  = MARGIN_LEFT + col * (LABEL_W + H_GAP)
        ly  = ph - MARGIN_TOP - (row + 1) * LABEL_H - row * V_GAP
        draw_label(c, lx, ly, v)
        page_slot += 1
    c.save()
    buf.seek(0)
    return buf.getvalue()

# ── Download bar (top) ────────────────────────────────────────
if st.session_state.lbl_pdf_ready:
    st.markdown('<div class="download-bar"><div class="download-bar-title">✅ Labels ready — print on Avery 8460 sheets</div></div>', unsafe_allow_html=True)
    st.download_button("📄 Download Label PDF", data=st.session_state.lbl_pdf_bytes,
                       file_name="loaner_labels.pdf", mime="application/pdf",
                       use_container_width=True, key="top_pdf")
    st.divider()

# ── Upload ────────────────────────────────────────────────────
uploaded = st.file_uploader("Upload vehicle list (CSV or Excel)", type=['csv','xlsx','xls'])
st.markdown('<p class="upload-hint">Columns needed: Unit Number · VIN · License Plate · Year · Make · Model</p>', unsafe_allow_html=True)

if uploaded:
    vehicles = parse_file(uploaded)
    st.session_state.lbl_vehicles  = vehicles
    st.session_state.lbl_pdf_ready = False

if st.session_state.lbl_vehicles:
    vehicles = st.session_state.lbl_vehicles
    count    = len(vehicles)

    st.success(f"✅ {count} vehicle{'s' if count != 1 else ''} loaded")

    pages_needed = max(1, -(-(st.session_state.lbl_start + count) // SLOTS))
    c1, c2 = st.columns(2)
    with c1: st.markdown(f'<div class="stat-box"><div class="stat-number">{count}</div><div class="stat-label">Vehicles</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="stat-box"><div class="stat-number">{pages_needed}</div><div class="stat-label">Pages needed</div></div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("### Choose start slot")
    st.markdown("Click a slot on the sheet to set where printing begins. Slots before it will be left blank.")

    # ── Interactive picker ────────────────────────────────────
    picker_html = f"""<!DOCTYPE html><html><head><style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:system-ui,sans-serif;}}
body{{background:transparent;padding:8px 0;}}
.controls{{display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap;}}
label{{font-size:13px;color:#555;}}
select{{font-size:13px;padding:4px 8px;border-radius:6px;border:1px solid #ccc;background:white;color:#111;}}
.wrap{{display:flex;gap:20px;align-items:flex-start;flex-wrap:wrap;}}
.sheet{{display:grid;grid-template-columns:repeat(3,76px);grid-template-rows:repeat(10,34px);gap:4px;}}
.lbl{{border:1px solid #ddd;border-radius:3px;display:flex;align-items:center;justify-content:center;
      font-size:11px;font-weight:500;background:#fff;color:#aaa;cursor:pointer;
      transition:transform 0.08s;user-select:none;}}
.lbl:hover{{transform:scale(1.06);}}
.lbl.skip  {{background:#f5f5f5;color:#ccc;}}
.lbl.start {{background:#1D9E75;color:white;border-color:#1D9E75;}}
.lbl.filled{{background:#1a3a5c;color:white;border-color:#1a3a5c;}}
.lbl.blank {{color:#eee;}}
.legend{{display:flex;flex-direction:column;gap:8px;}}
.leg{{display:flex;align-items:center;gap:8px;font-size:12px;color:#555;}}
.lb{{width:16px;height:16px;border-radius:3px;flex-shrink:0;border:1px solid #ddd;}}
.hint{{font-size:11px;color:#999;margin-top:6px;}}
</style></head><body>
<div class="controls">
  <label>Start slot:</label>
  <select id="dd" onchange="setStart(parseInt(this.value))"></select>
</div>
<div class="wrap">
  <div>
    <div style="font-size:11px;color:#999;margin-bottom:6px;">Click any slot to set start</div>
    <div class="sheet" id="sheet"></div>
    <div class="hint">→ left to right &nbsp;|&nbsp; ↓ top to bottom</div>
  </div>
  <div class="legend">
    <div style="font-size:12px;font-weight:600;color:#333;margin-bottom:2px;">Legend</div>
    <div class="leg"><div class="lb" style="background:#1D9E75;border-color:#1D9E75;"></div>Start position</div>
    <div class="leg"><div class="lb" style="background:#1a3a5c;border-color:#1a3a5c;"></div>Printed label</div>
    <div class="leg"><div class="lb" style="background:#f5f5f5;"></div>Skipped (used)</div>
    <div class="leg"><div class="lb"></div>Blank</div>
    <div style="margin-top:8px;font-size:11px;color:#888;max-width:130px;line-height:1.5;">{count} label{'s' if count!=1 else ''} to print.</div>
  </div>
</div>
<script>
const COLS=3,ROWS=10,SLOTS=COLS*ROWS,COUNT={count};
let cur=0;
function setStart(s){{
  cur=s;
  document.getElementById('dd').value=s;
  render();
  window.parent.postMessage({{type:'streamlit:setComponentValue',value:s}},'*');
}}
function render(){{
  const sheet=document.getElementById('sheet');
  sheet.innerHTML='';
  for(let i=0;i<SLOTS;i++){{
    const d=document.createElement('div');
    const pi=i-cur;
    if(i<cur){{d.className='lbl skip';d.textContent='';}}
    else if(pi===0){{d.className='lbl start';d.textContent='1';}}
    else if(pi>0&&pi<COUNT){{d.className='lbl filled';d.textContent=pi+1;}}
    else{{d.className='lbl blank';d.textContent='';}}
    d.addEventListener('click',()=>setStart(i));
    sheet.appendChild(d);
  }}
}}
function init(){{
  const sel=document.getElementById('dd');
  for(let i=0;i<SLOTS;i++){{
    const r=Math.floor(i/COLS)+1,c=(i%COLS)+1;
    const cn=c===1?'left':c===2?'center':'right';
    const o=document.createElement('option');
    o.value=i;o.textContent='Slot '+(i+1)+' (row '+r+', '+cn+')';
    sel.appendChild(o);
  }}
  render();
}}
init();
</script></body></html>"""

    components.html(picker_html, height=400, scrolling=False)

    st.markdown("**Confirm start slot** (1 = top-left corner):")
    slot_num = st.number_input(
        "start_slot", min_value=1, max_value=SLOTS,
        value=st.session_state.lbl_start + 1,
        step=1, label_visibility="collapsed"
    )
    st.session_state.lbl_start = int(slot_num) - 1

    pages_needed = max(1, -(-(st.session_state.lbl_start + count) // SLOTS))
    st.caption(f"Starting at slot {st.session_state.lbl_start + 1} · {count} labels · **{pages_needed} page{'s' if pages_needed != 1 else ''}**")

    st.divider()

    if st.button("🖨️ GENERATE LABELS PDF"):
        with st.spinner("Generating labels…"):
            pdf_bytes = generate_pdf(vehicles, st.session_state.lbl_start)
            st.session_state.lbl_pdf_bytes = pdf_bytes
            st.session_state.lbl_pdf_ready = True
        st.rerun()
