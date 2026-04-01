# ──────────────────────────────────────────────────────────────────
# LEASE PDF EXTRACTOR v2 — coordinate-based word extraction
# ──────────────────────────────────────────────────────────────────
!pip install pdfplumber openpyxl -q

import pdfplumber, re, os
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from google.colab import files

# ── Constants ─────────────────────────────────────────────────────
MONTH_NAMES = {'january','february','march','april','may','june','july',
               'august','september','october','november','december'}
ORDINAL_RE  = re.compile(r'^\d{1,2}(?:st|nd|rd|th)$', re.I)
YEAR_RE     = re.compile(r'^(20\d{2}|19\d{2})$')
AMOUNT_RE   = re.compile(r'^\d{1,6}[.,]\d{2}$')

RIGHT_X0, RIGHT_X1 = 308, 576   # right column x-bounds
LEFT_X0,  LEFT_X1  = 48,  305   # left column x-bounds
DEBUG = False                    # set True to see word dumps

# ── Word helpers ──────────────────────────────────────────────────

def get_words(page, x0=0, top=0, x1=None, bottom=None, tol=2):
    """Extract words inside a bounding box (with small tolerance)."""
    if x1 is None:     x1     = page.width
    if bottom is None: bottom = page.height
    return [w for w in page.extract_words(x_tolerance=3, y_tolerance=3)
            if w['x0'] >= x0 - tol and w['x1'] <= x1 + tol
            and w['top'] >= top - tol and w['bottom'] <= bottom + tol]

def at_y(words, y_center, y_tol=8):
    """Return words within y_tol of y_center."""
    return [w for w in words if abs(w['top'] - y_center) <= y_tol]

def classify(text):
    t = text.strip().rstrip(',.')
    if t.lower() in MONTH_NAMES:       return 'month'
    if ORDINAL_RE.match(t):            return 'ordinal'
    if YEAR_RE.match(t):               return 'year'
    if AMOUNT_RE.match(t.replace(',','')): return 'amount'
    return 'other'

# ── Field extractors ──────────────────────────────────────────────

def get_date_of_contract(pages):
    """
    'June 2, 2025' — in the header strip, left side, top ≈ 70-80.
    Words: June(x0≈180), 2,(x0≈208), 2025(x0≈225), all at top≈70.9
    """
    page  = pages[0]
    words = get_words(page, x0=120, top=58, x1=310, bottom=86)
    if DEBUG: print(f"[DATE] {[(w['text'], round(w['top'],1)) for w in words]}")

    # Find the month word → gather its whole row
    for w in sorted(words, key=lambda x: x['x0']):
        if classify(w['text']) == 'month':
            row = sorted(at_y(words, w['top'], y_tol=6), key=lambda x: x['x0'])
            return ' '.join(r['text'] for r in row).strip()

    # Fallback: crop + regex
    t = (page.crop((0, 58, 400, 90)).extract_text(x_tolerance=3, y_tolerance=3) or '')
    m = re.search(r'([A-Za-z]+ \d{1,2},?\s*\d{4})', t)
    return m.group(1).strip() if m else "Not found"


def get_parties(pages):
    """
    Resident names sit on the line AFTER 'Lease Contract):' (top≈146.8),
    so they appear at top≈154-165 in the left column.
    Words confirmed: Luis(59.8), Monge(88.2), Hernandez,(122.2),
                     Diana(184.6), Hernandez(218.6), Dias(275.3) — all top=154.8
    """
    page = pages[0]
    SKIP = {'lease','contract','list','all','people','signing','the','is',
            'between','you,','you','resident','residents','"lease")',
            'lease")', '(list','(sometimes','referred','to','as'}

    for y0, y1 in [(148, 168), (144, 178), (140, 190)]:
        words = get_words(page, x0=LEFT_X0, top=y0, x1=LEFT_X1, bottom=y1)
        names = [w for w in words
                 if w['text'].lower().strip('():,"') not in SKIP
                 and len(w['text'].strip('():," ')) > 1
                 and not w['text'].startswith('"')]
        if DEBUG: print(f"[PARTIES y={y0}-{y1}] {[w['text'] for w in names]}")
        if names:
            names.sort(key=lambda w: (round(w['top']), w['x0']))
            result = ' '.join(w['text'] for w in names)
            if re.search(r'[A-Z][a-z]{2,}', result):   # looks like a real name
                return result.strip()

    return "Not found"


def get_lease_term(pages):
    """
    Confirmed word positions in right column:
      '12th'   x0=524.1  top=134.3   ← begin day ordinal
      'June'   x0=369.9  top=144.4   ← begin month
      '2025'   x0=440.8  top=144.4   ← begin year
      '16th'   x0=344.1  top=155.3   ← end day ordinal
      'August' x0=429.7  top=154.8   ← end month
      '2026'   x0=506.1  top=155.3   ← end year
    Strategy: collect all ordinals/months/years in right col, y=120-180,
    then assign by sort order.
    """
    for pg in pages:
        words = get_words(pg, x0=RIGHT_X0, top=118, x1=RIGHT_X1, bottom=178)
        if not words:
            continue

        ordinals = sorted([w for w in words if classify(w['text']) == 'ordinal'],
                          key=lambda w: w['top'])
        months   = sorted([w for w in words if classify(w['text']) == 'month'],
                          key=lambda w: w['top'])
        years    = sorted([w for w in words if classify(w['text']) == 'year'],
                          key=lambda w: w['top'])

        if DEBUG:
            print(f"[LEASE TERM] ord={[w['text'] for w in ordinals]}, "
                  f"mon={[w['text'] for w in months]}, yr={[w['text'] for w in years]}")

        if ordinals and months and years:
            bm = months[0]['text'].rstrip(',')
            by = years[0]['text']
            bd = ordinals[0]['text']
            begin = f"{bm} {bd}, {by}"

            em = months[-1]['text'].rstrip(',')
            ey = years[-1]['text']
            ed = ordinals[1]['text'] if len(ordinals) > 1 else ordinals[0]['text']
            end = f"{em} {ed}, {ey}"

            return begin, end

    return "Not found", "Not found"


def get_security_deposit(pages):
    """
    Confirmed: '295.00' at x0=454.1, top=587.8 in right column.
    Form anchor: 'residents in the apartment is $' at top≈590.
    """
    for pg in pages:
        words = get_words(pg, x0=RIGHT_X0, top=555, x1=RIGHT_X1, bottom=625)
        if DEBUG: print(f"[DEPOSIT] {[(w['text'], round(w['top'],1)) for w in words]}")

        for w in sorted(words, key=lambda x: x['top']):
            val = w['text'].replace(',', '')
            if AMOUNT_RE.match(val):
                # Confirm: anchor words on same y-band
                ctx = ' '.join(nw['text'].lower() for nw in at_y(words, w['top'], y_tol=12))
                if any(k in ctx for k in ('residents', '$', 'apartment', 'deposit')):
                    return val

        # Relaxed — any amount in a wider band
        words2 = get_words(pg, x0=RIGHT_X0, top=540, x1=RIGHT_X1, bottom=650)
        amounts = [(w, w['top']) for w in words2
                   if AMOUNT_RE.match(w['text'].replace(',', ''))]
        if amounts:
            amounts.sort(key=lambda x: x[1])
            for w, t in amounts:
                if 565 <= t <= 615:
                    return w['text'].replace(',', '')
            return amounts[0][0]['text'].replace(',', '')

    return "Not found"


def get_monthly_rent(pages):
    """
    Section 6 (RENT AND CHARGES) is on page index 1 of the original PDF.
    Fallback: Use liquidated damages field (= 1 month's rent) confirmed at
              right col, top≈263.4, which the form states equals one month's rent.
    """
    RENT_KW = ['rent and charges', 'rent and charge']

    for pg in pages:
        all_w  = pg.extract_words(x_tolerance=3, y_tolerance=3)
        full_t = ' '.join(w['text'].lower() for w in all_w)
        if not any(kw in full_t for kw in RENT_KW):
            continue

        if DEBUG: print("[RENT] Found 'RENT AND CHARGES' section")

        # Strategy A: find amount on same line as "per month"
        for w in all_w:
            if w['text'].lower().rstrip('.,') == 'month':
                row = at_y(all_w, w['top'], y_tol=10)
                row_text = ' '.join(rw['text'].lower()
                                    for rw in sorted(row, key=lambda x: x['x0']))
                if 'per' in row_text:
                    for rw in sorted(row, key=lambda x: x['x0']):
                        val = rw['text'].replace(',', '').lstrip('$')
                        if AMOUNT_RE.match(val):
                            return val

        # Strategy B: find $ sign then next amount word
        for w in all_w:
            if w['text'] == '$':
                after = sorted([nw for nw in all_w
                                if abs(nw['top'] - w['top']) < 8 and nw['x0'] > w['x0']],
                               key=lambda x: x['x0'])
                for nw in after:
                    val = nw['text'].replace(',', '')
                    if AMOUNT_RE.match(val):
                        return val

    # Fallback: liquidated damages in Section 3 right col top≈255-275
    # (Form states this equals "one month's rent")
    for pg in pages:
        words = get_words(pg, x0=RIGHT_X0, top=253, x1=RIGHT_X1, bottom=278)
        for w in sorted(words, key=lambda x: x['x0']):
            val = w['text'].replace(',', '')
            if AMOUNT_RE.match(val):
                return val   # liquidated damages = 1 month rent per the form

    return "Not found"

# ── Orchestrator ──────────────────────────────────────────────────

def process_pdf(pdf_path):
    fname = os.path.basename(pdf_path)
    print(f"\n{'─'*58}")
    print(f"  📄  {fname}")
    print(f"{'─'*58}")

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        n = min(4, total)
        print(f"  Pages in PDF: {total}  |  Scanning first {n}\n")
        pages = [pdf.pages[i] for i in range(n)]

        date       = get_date_of_contract(pages)
        parties    = get_parties(pages)
        begin, end = get_lease_term(pages)
        deposit    = get_security_deposit(pages)
        rent       = get_monthly_rent(pages)

    row = {
        'File Name':              fname,
        'Date of Lease Contract': date,
        'Parties (Residents)':    parties,
        'Lease Begin Date':       begin,
        'Lease End Date':         end,
        'Security Deposit ($)':   deposit,
        'Monthly Rent ($)':       rent,
    }
    pad = max(len(k) for k in row)
    for k, v in row.items():
        icon = "✅" if v not in ("Not found", "") else "⚠️ "
        print(f"  {icon}  {k:<{pad}}  →  {v}")
    return row

# ── Excel export ──────────────────────────────────────────────────

def to_excel(rows, path='lease_data_extracted.xlsx'):
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Lease Data')
        ws = writer.sheets['Lease Data']
        fill = PatternFill('solid', fgColor='1F4E79')
        font = Font(bold=True, color='FFFFFF', size=11)
        for cell in ws[1]:
            cell.fill = fill; cell.font = font
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 32
        for col in ws.columns:
            w = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(w + 4, 60)
        ws.freeze_panes = 'A2'
    print(f"\n  ✅  Excel saved → {path}")
    return path

# ── Run ───────────────────────────────────────────────────────────

print("📁  Upload one or more lease PDF files…")
uploaded = files.upload()

rows = []
for fname, content in uploaded.items():
    tmp = f'/tmp/{fname}'
    with open(tmp, 'wb') as f:
        f.write(content)
    try:
        rows.append(process_pdf(tmp))
    except Exception:
        import traceback; traceback.print_exc()
        rows.append({'File Name': fname,
                     **{k: 'ERROR' for k in [
                         'Date of Lease Contract', 'Parties (Residents)',
                         'Lease Begin Date', 'Lease End Date',
                         'Security Deposit ($)', 'Monthly Rent ($)']}})

if rows:
    out = to_excel(rows)
    files.download(out)
    print("\n🎉  Done!")
