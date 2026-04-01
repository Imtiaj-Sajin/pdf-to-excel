!pip install -q pdfplumber pandas openpyxl

import pdfplumber
import pandas as pd
import re

# --- Hardcoded file path as requested ---
pdf_path = "/content/1-101.pdf" 

def extract_visual_text(crop_obj, y_tol=4.5):
    """
    Extracts words using their absolute X/Y coordinates on the page.
    If floating form-fill text (like a date or amount) shares the same 
    vertical Y-coordinate as the document text, it stitches them together 
    left-to-right. This reconstructs the visual sentences perfectly.
    """
    words = crop_obj.extract_words(keep_blank_chars=False, x_tolerance=3, y_tolerance=3)
    if not words: 
        return ""
    
    # Sort words by vertical position (top) first, then horizontal (x0)
    words.sort(key=lambda w: (w['top'], w['x0']))
    
    lines = []
    current_line = [words[0]]
    
    for w in words[1:]:
        # Calculate average Y position of the current line being built
        avg_top = sum(cw['top'] for cw in current_line) / len(current_line)
        
        # If the word is on the same visual line (within tolerance)
        if abs(w['top'] - avg_top) <= y_tol:
            current_line.append(w)
        else:
            # Sort the completed line left-to-right before saving
            current_line.sort(key=lambda cw: cw['x0'])
            lines.append(" ".join([cw['text'] for cw in current_line]))
            current_line = [w]
            
    if current_line:
        current_line.sort(key=lambda cw: cw['x0'])
        lines.append(" ".join([cw['text'] for cw in current_line]))
        
    return "\n".join(lines)

def process_lease_pdf(file_path):
    data = {
        "File Name": file_path.split("/")[-1],
        "Lease Date": "Not Found",
        "Parties/Residents": "Not Found",
        "Lease Begin Date": "Not Found",
        "Lease End Date": "Not Found",
        "Security Deposit ($)": "Not Found",
        "Rent Amount ($)": "Not Found"
    }

    try:
        with pdfplumber.open(file_path) as pdf:
            # === PAGE 1 EXTRACTION ===
            p1 = pdf.pages[0]
            w, h = p1.width, p1.height
            mid_x = w / 2
            
            # Crop 1: Top Header (y=0 to 120)
            header_crop = p1.within_bbox((0, 0, w, 120))
            header_text = extract_visual_text(header_crop)
            
            # Crop 2: Left Column (y=120 to end)
            left_crop = p1.within_bbox((0, 120, mid_x + 5, h))
            left_text = extract_visual_text(left_crop)
            
            # Crop 3: Right Column (y=120 to end)
            right_crop = p1.within_bbox((mid_x - 5, 120, w, h))
            right_text = extract_visual_text(right_crop)
            
            # === PAGE 2 EXTRACTION (For Rent) ===
            p2_text_flat = ""
            if len(pdf.pages) > 1:
                p2 = pdf.pages[1]
                p2_text_flat = extract_visual_text(p2).replace('\n', ' ')

            # ---------------------------------------------------------
            # DATA EXTRACTION USING THE RECONSTRUCTED VISUAL TEXT
            # ---------------------------------------------------------

            # 1. Lease Date (Header)
            m_date = re.search(r"Date of Lease Contract:\s*(.*)", header_text, re.IGNORECASE)
            if m_date:
                data["Lease Date"] = m_date.group(1).replace(" (when the Lease Contract is filled out)", "").strip()

            # 2. Parties/Residents (Left Column)
            # Looks for the line directly beneath "Lease Contract):"
            m_parties = re.search(r"Lease Contract\):\s*\n([A-Za-z\s,]+)\nand us, the owner", left_text)
            if m_parties:
                data["Parties/Residents"] = m_parties.group(1).strip()
            else:
                # Fallback if they merged onto one line
                m_alt = re.search(r"Lease Contract\):\s*(.*?)and us, the owner", left_text.replace('\n', ' '))
                if m_alt:
                    data["Parties/Residents"] = m_alt.group(1).strip()

            # Flatten right column text to safely jump across newlines
            right_flat = right_text.replace('\n', ' ')

            # 3. Lease Begin/End Dates (Right Column)
            m_begin = re.search(r"begins on the\s*(.*?)\s*day\s*of\s*(.*?)\s*,", right_flat, re.IGNORECASE)
            if m_begin:
                data["Lease Begin Date"] = f"{m_begin.group(1).strip()} {m_begin.group(2).strip()}"

            m_end = re.search(r"ends at.*?the\s*(.*?)\s*day\s*of\s*(.*?)\s*\.", right_flat, re.IGNORECASE)
            if m_end:
                data["Lease End Date"] = f"{m_end.group(1).strip()} {m_end.group(2).strip()}"

            # 4. Security Deposit (Right Column)
            m_sec = re.search(r"SECURITY DEPOSIT.*?is\s*\$\s*([\d,]+\.\d{2})", right_flat, re.IGNORECASE)
            if m_sec:
                data["Security Deposit ($)"] = m_sec.group(1).strip()

            # 5. Rent Amount (Page 2)
            m_rent = re.search(r"RENT AND CHARGES.*?pay\s*\$\s*([\d,]+\.\d{2})", p2_text_flat, re.IGNORECASE)
            if m_rent:
                data["Rent Amount ($)"] = m_rent.group(1).strip()

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

    return data

# ==========================================
# MAIN EXECUTION
# ==========================================
print(f"Processing target file: {pdf_path} ...")

# Run the extraction
extracted_info = process_lease_pdf(pdf_path)

# Convert to Pandas DataFrame
df = pd.DataFrame([extracted_info])

# Export to Excel
excel_filename = "extracted_lease_data.xlsx"
df.to_excel(excel_filename, index=False)

print("\n--- EXTRACTION RESULTS ---")
for key, value in extracted_info.items():
    print(f"{key}: {value}")

print(f"\nSuccessfully generated '{excel_filename}'.")
