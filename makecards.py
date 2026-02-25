import math
import os
import pandas as pd
import re
import sys
import tkinter as tk

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.platypus import Preformatted
from reportlab.lib.enums import TA_LEFT

from tkinter import filedialog, messagebox

# Written by Tyrone Rees for Wantage White Horses Swim Club
# itoffice@wwhsc.co.uk
# 
# CSV files inputted are to be in the format exported from 
# Meet Manager's 'Lane/Timer sheets' report

def script_folder_default():
    # preferred: folder where the script file lives
    try:
        return os.path.dirname(os.path.abspath(__file__)) or "."
    except NameError:
        # __file__ may not exist in interactive environments; fall back to cwd
        return os.getcwd()

def ask_for_csv_file(title="Select CSV file"):
    root = tk.Tk()
    root.withdraw()  # hide main window
    # show a simple prompt first
    messagebox.showinfo(title="Choose CSV", message="Please select the CSV file that contains the swim entries.")
    
    root.attributes('-topmost', True)  # bring dialog to front
    filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
    initial_dir = script_folder_default()
    path = filedialog.askopenfilename(title=title, initialdir=initial_dir or ".", filetypes=filetypes)
    root.destroy()
    return path

# CONFIG
#INPUT_CSV = "wwhsc100m2025.csv"   # update if needed
# Call the dialog to get the CSV path
INPUT_CSV = ask_for_csv_file()
if not INPUT_CSV:
    print("No file selected. Exiting.")
    sys.exit(0)

OUTPUT_PDF = os.path.splitext(INPUT_CSV)[0] + "_time_keepers.pdf"
JUDGE_OUTPUT_PDF = os.path.splitext(INPUT_CSV)[0] + "_judge_slip.pdf"
REGISTER_OUTPUT_PDF = os.path.splitext(INPUT_CSV)[0] + "_register.pdf"

# Which columns to extract (0-based): 1 = Event, 3 = Lane, 4 = Heat, 5 = Name
COL_INDEXES = [1, 3, 4, 5]

# Page/card layout: 4 columns x 2 rows = 8 cards per A4 portrait page
COLUMNS = 2
ROWS = 3
CARDS_PER_PAGE = COLUMNS * ROWS

JUDGE_COLUMNS = 2
JUDGE_ROWS = 2
JUDGE_CARDS_PER_PAGE = JUDGE_COLUMNS * JUDGE_ROWS

# visual settings
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN_MM = 5             # page margin in mm
CARD_GAP_MM = 0           # gap between cards in mm
FONT_NAME = "Helvetica"
TITLE_FONT_SIZE = 12
FIELD_FONT_SIZE = 12
LINE_SPACING = 4

# helper to clean field strings
def clean_field(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # treat values like single-space or empty same as blank
    if s == "":
        return ""
    return s

# remove trailing sex/age tokens from name strings
AGE_SUFFIX_RE = re.compile(r'\s+(?:[MWF]\s*)?\d{1,3}$', re.IGNORECASE)

# clean and normalize Event strings
DISTANCE_AND_STROKE_RE = re.compile(r'(\d{2,3})\s*(?:SC\s*)?Meter\s*([A-Za-z]+)', re.IGNORECASE)
EVENT_NUMBER_RE = re.compile(r'(Event\s*\d+)', re.IGNORECASE)

def clean_name_field(x):
    s = clean_field(x)
    if s == "":
        return ""
    # remove obvious multiple spaces and leading/trailing spaces first
    s = re.sub(r'\s+', ' ', s).strip()
    # remove suffix like " M11", "W12", " 9", " F 10"
    s = AGE_SUFFIX_RE.sub('', s).strip()
    return s

def clean_event_field(s):
    s = clean_field(s)
    if not s:
        return ""
    # try to extract "Event N"
    m_evt = EVENT_NUMBER_RE.search(s)
    evt = m_evt.group(1).title().replace("  ", " ") if m_evt else ""

    # Extract gender, including compound gender labels like "Open/Male", "Male/Female"
    m_gender = re.search(r'\b(?:Mixed|Male|Female|Open)(?:\s*/\s*(?:Mixed|Male|Female|Open))*\b', s, flags=re.IGNORECASE)
    gender = m_gender.group(0).title().replace("/", "/") if m_gender else ""


    # try to find distance and stroke (e.g., "200 SC Meter Butterfly")
    m_ds = DISTANCE_AND_STROKE_RE.search(s)
    if m_ds:
        distance = m_ds.group(1)
        stroke = m_ds.group(2).strip().capitalize()
        # ensure "Meter" singular form and spacing
        core = f"{distance}M {stroke}"
        return " ".join(filter(None, [evt, gender, core]))
        #return (evt + " " + core).strip() if evt else core

    # fallback: remove gender words, ages, SC and Finals, then collapse spaces
    s2 = re.sub(r'\bSC\b', '', s2, flags=re.IGNORECASE)                         # remove SC
    s2 = re.sub(r'\bFinals\b', '', s2, flags=re.IGNORECASE)                     # remove Finals
    # remove extra punctuation and collapse whitespace
    s2 = re.sub(r'[\/,]+', ' ', s2)
    s2 = re.sub(r'\s+', ' ', s2).strip()
    # if we still have something like "Event 1 200 Butterfly", ensure "Meter" inserted if distance present without "Meter"
    m_simple = re.search(r'(\d{2,3})\s+([A-Za-z]+)', s2)
    if m_simple:
        distance = m_simple.group(1)
        stroke = m_simple.group(2).capitalize()
        core = f"str = {distance}M {stroke}"
        return " ".join(filter(None, [evt, gender, core]))

    # Final fallback: just return cleaned string with evt and gender
    return " ".join(filter(None, [evt, gender, s2]))


def load_and_extract(path):
    # read CSV without forcing header - file has many quoted columns
    df = pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
    # extract requested columns and clean
    #extracted = df.iloc[:, COL_INDEXES].applymap(clean_field)
    extracted = df.iloc[:, COL_INDEXES].apply(lambda col: col.map(clean_field))
    extracted.columns = ["Event", "Lane", "Heat", "Name"]
    # Apply the event cleaning
    extracted["Event"] = extracted["Event"].apply(clean_event_field)
    # clean names specifically
    extracted["Name"] = extracted["Name"].apply(clean_name_field)
    # Filter rows where Lane is non-empty (ignore the even/blank lines)
    extracted = extracted[extracted["Lane"].str.strip() != ""].reset_index(drop=True)
    return extracted

def draw_card(c, x, y, w, h, data):
    # c: canvas, (x,y) bottom-left of card, w/h in points
    # draw border
    c.setStrokeColor(colors.black)
    c.rect(x, y, w, h, stroke=1, fill=0)

    # padding inside card
    pad = 2 * mm
    tx = x + pad
    ty = y + h - pad
    
    # Title: Lane (allow wrapping)
    c.setFont(FONT_NAME, TITLE_FONT_SIZE)
    # we use Paragraph for wrapping the event text
    style = ParagraphStyle("card_title", 
			fontName=FONT_NAME,
			fontSize=TITLE_FONT_SIZE,
			leading=TITLE_FONT_SIZE + 2,
			alignment=TA_LEFT)
    p = Paragraph("Time Keepers", style)
    w_available = w - 2 * pad
    # draw paragraph at (tx, ty) where platypus uses top-left; create a frame-like draw
    w_needed, h_needed = p.wrap(w_available, h)
    p.drawOn(c, tx, ty - h_needed)
    style2 = ParagraphStyle("lane",
			fontName=FONT_NAME,
			fontSize=TITLE_FONT_SIZE,
			leading=TITLE_FONT_SIZE + 2,
			alignment=TA_LEFT)
    p = Paragraph(data['Lane'], style)
    w_available = w - 2 * pad
    # draw paragraph at (tx, ty) where platypus uses top-left; create a frame-like draw
    w_needed, h_needed = p.wrap(w_available, h)
    p.drawOn(c, tx + 81*mm, ty - h_needed)
    cursor_y = ty - h_needed - (LINE_SPACING)

    # Leave a blank line	
    cursor_y = cursor_y - h_needed - (LINE_SPACING)

    # Remaining fields: Event, Heat, Name
    c.setFont(FONT_NAME, FIELD_FONT_SIZE)
   
    # Name (make it larger)
    name_style = ParagraphStyle("name_style", 
				fontName=FONT_NAME, 
				fontSize=FIELD_FONT_SIZE, 
				leading=FIELD_FONT_SIZE + 2, 
				alignment=TA_LEFT)
    name_text = f"<b>{data['Name']}</b>"
    p_name = Paragraph(name_text, name_style)
    w_n, h_n = p_name.wrap(w_available, h)
    p_name.drawOn(c, tx, cursor_y - h_n)
    cursor_y = cursor_y - h_n - (LINE_SPACING)

    # Event
    event_words = data['Event'].replace("&", "&amp;").split()
    event_number = " ".join(event_words[:2])
    event_description = " ".join(event_words[2:])
    
    event_text = f"<b>{event_number}</b> {event_description}"
    p_event = Paragraph(event_text, 
			ParagraphStyle("event_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_l, h_l = p_event.wrap(w_available, h)
    p_event.drawOn(c, tx, cursor_y - h_l)
    cursor_y = cursor_y - h_l - (LINE_SPACING)

    # Heat
    heat_text = f"<b>{data['Heat']}</b> "
    p_heat = Paragraph(heat_text, 
			ParagraphStyle("heat_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_h, h_h = p_heat.wrap(w_available, h)
    p_heat.drawOn(c, tx, cursor_y - h_h)
    cursor_y = cursor_y - h_h - (LINE_SPACING)

    cursor_y = cursor_y - 5 * mm
 
    p_time = Paragraph("Time:", 
			ParagraphStyle("time_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_t, h_t = p_time.wrap(w_available, h)
    p_time.drawOn(c, tx, cursor_y - h_t)
    cursor_y = cursor_y - h_t - (LINE_SPACING)

    p_time = Preformatted("Min                       Sec                       Tenths", 
			ParagraphStyle("time_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_t, h_t = p_time.wrap(w_available, h)
    p_time.drawOn(c, tx, cursor_y - h_t)
    cursor_y = cursor_y - h_t - (LINE_SPACING)

	
    # end draw

def draw_judge_card(c, x, y, w, h, data):
    # c: canvas, (x,y) bottom-left of card, w/h in points
    # draw border
    c.setStrokeColor(colors.black)
    c.rect(x, y, w, h, stroke=1, fill=0)

    # padding inside card
    pad = 10 * mm
    tx = x + pad
    ty = y + h - pad
    
    # Title: Lane (allow wrapping)
    c.setFont(FONT_NAME, TITLE_FONT_SIZE)
    # we use Paragraph for wrapping the event text
    style = ParagraphStyle("card_title", 
			fontName=FONT_NAME,
			fontSize=TITLE_FONT_SIZE+2,
			leading=TITLE_FONT_SIZE + 2,
			alignment=TA_LEFT)
    p = Paragraph("Judge Place Slip", style)
    w_available = w - 2 * pad
    # draw paragraph at (tx, ty) where platypus uses top-left; create a frame-like draw
    w_needed, h_needed = p.wrap(w_available, h)
    p.drawOn(c, tx, ty - h_needed)
    cursor_y = ty - h_needed - (LINE_SPACING)

    # Leave a blank line	
    cursor_y = cursor_y - h_needed - (LINE_SPACING)

    # Remaining fields: Event, Heat
    c.setFont(FONT_NAME, FIELD_FONT_SIZE)

    # Event
    event_words = data['Event'].replace("&", "&amp;").split()
    event_number = " ".join(event_words[:2])
    event_description = " ".join(event_words[2:])
    
    event_number_text = f"<b>{event_number}</b>"
    p_event = Paragraph(event_number_text, 
			ParagraphStyle("event_number", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE+2, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_l, h_l = p_event.wrap(w_available, h)
    p_event.drawOn(c, tx, cursor_y - h_l)
    cursor_y = cursor_y - h_l - (LINE_SPACING)

    event_description_text = f"{event_description}"
    p_event = Paragraph(event_description_text, 
			ParagraphStyle("event_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE+2, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_l, h_l = p_event.wrap(w_available, h)
    p_event.drawOn(c, tx, cursor_y - h_l)
    cursor_y = cursor_y - h_l - (LINE_SPACING)

    cursor_y = cursor_y - 0.5 * h_needed - (LINE_SPACING)
       
       
    
    # Heat
    heat_text = f"<b>{data['Heat']}</b> "
    p_heat = Paragraph(heat_text, 
			ParagraphStyle("heat_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE+2, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_h, h_h = p_heat.wrap(w_available, h)
    p_heat.drawOn(c, tx, cursor_y - h_h)
    cursor_y = cursor_y - h_h - (LINE_SPACING)

    cursor_y = cursor_y - 5 * mm
 

    p_time = Preformatted("Pos                  Lane         DQ", 
			ParagraphStyle("time_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE+2, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
    w_t, h_t = p_time.wrap(w_available, h)
    p_time.drawOn(c, tx, cursor_y - h_t)
    cursor_y = cursor_y - h_t - (LINE_SPACING)
    for i in range(1,7):
       cursor_y = cursor_y - 0.35 * h_needed - (LINE_SPACING)
       
       p_time = Preformatted(f"{i}", 
			ParagraphStyle("time_style", 
					fontName=FONT_NAME, 
					fontSize=FIELD_FONT_SIZE+4, 
					leading=FIELD_FONT_SIZE + 2, 
					alignment=TA_LEFT))
       w_t, h_t = p_time.wrap(w_available, h)
       p_time.drawOn(c, tx, cursor_y - h_t)
       cursor_y = cursor_y - h_t - (LINE_SPACING)
        

	
    # end draw


def make_pdf(df, outpath):
    c = canvas.Canvas(outpath, pagesize=A4)
    margin = MARGIN_MM * mm
    gap = CARD_GAP_MM * mm

    # compute card size
    usable_w = PAGE_WIDTH - 2 * margin - (COLUMNS - 1) * gap
    usable_h = PAGE_HEIGHT - 2 * margin - (ROWS - 1) * gap
    card_w = usable_w / COLUMNS
    card_h = usable_h / ROWS

    total = len(df)
    pages = math.ceil(total / CARDS_PER_PAGE)
    idx = 0

    for p in range(pages):
        for r in range(ROWS):
            for col in range(COLUMNS):
                if idx >= total:
                    break
                # compute bottom-left coords of this card
                x = margin + col * (card_w + gap)
                # rows indexed top-to-bottom, but y is from bottom, so compute accordingly
                y = margin + (ROWS - 1 - r) * (card_h + gap)
                row = df.iloc[idx]
                data = {
                    "Event": row["Event"],
                    "Lane": row["Lane"],
                    "Heat": row["Heat"],
                    "Name": row["Name"]
                }
                draw_card(c, x, y, card_w, card_h, data)
                idx += 1
        c.showPage()
    c.save()
    print(f"Saved {outpath} with {total} cards on {pages} pages.")

def make_judge_pdf(df, outpath):
    c = canvas.Canvas(outpath, pagesize=A4)
    margin = MARGIN_MM * mm
    gap = CARD_GAP_MM * mm

    # compute card size
    usable_w = PAGE_WIDTH - 2 * margin - (JUDGE_COLUMNS - 1) * gap
    usable_h = PAGE_HEIGHT - 2 * margin - (JUDGE_ROWS - 1) * gap
    card_w = usable_w / JUDGE_COLUMNS
    card_h = usable_h / JUDGE_ROWS

    #print(df)
    compressed_df = df[['Event', 'Heat']].drop_duplicates().reset_index(drop=True)
    #print(compressed_df)
    total = len(compressed_df)
    pages = math.ceil(total / JUDGE_CARDS_PER_PAGE)
    idx = 0
    
    for p in range(pages):
        for r in range(JUDGE_ROWS):
            for col in range(JUDGE_COLUMNS):
                if idx >= total:
                    break
                # compute bottom-left coords of this card
                x = margin + col * (card_w + gap)
                # rows indexed top-to-bottom, but y is from bottom, so compute accordingly
                y = margin + (JUDGE_ROWS - 1 - r) * (card_h + gap)
                row = df.iloc[idx]
                data = {
                    "Event": row["Event"],
                    "Heat": row["Heat"],
                }
                draw_judge_card(c, x, y, card_w, card_h, data)
                idx += 1
        c.showPage()
    c.save()
    print(f"Saved {outpath} with {total} cards on {pages} pages.")

def make_register_pdf(df, outpath):
    names = df['Name'][~df['Name'].str.contains(r'^_+$')].dropna().unique()
    names = sorted(names)

    raw_csv = pd.read_csv(INPUT_CSV, header=None, dtype=str, keep_default_na=False)
    event_title = raw_csv.iloc[0,0]

    c = canvas.Canvas(outpath, pagesize=A4)
    width, height = A4
    margin = 50
    line_height = 20
    y = height - margin

    checked_in_start = 250
    checked_out_start = 380

    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y, event_title)
    y -= line_height
    
    c.setFont("Helvetica", 14)
    c.drawString(margin, y, "Swimmer Check-In/Out Register")
    y -= line_height * 2

    c.drawString(margin + checked_in_start, y, "Checked in?")
    c.drawString(margin + checked_out_start, y, "Checked out?")
    y -= line_height
    for name in names:
        if y < margin:
            c.showPage()
            y = height - margin
            c.setFont("Helvetica", 14)

        c.drawString(margin + 10, y, name)
        c.rect(margin + checked_in_start + 30, y - 2.5, 12, 12)  # Check-in box
        c.rect(margin + checked_out_start + 35, y - 2.5, 12, 12)  # Check-out box
        line_y = y - 10
        c.line(margin, line_y, width-margin, line_y)
        
        y -= line_height + 5
    print(f"Saved {outpath} with swimmers' names")
    c.save()
    
    
if __name__ == "__main__":
    df = load_and_extract(INPUT_CSV)
    #print(df)
    if df.empty:
        print("No rows found with a Lane value. Check the CSV or the column indexes.")
    else:
        make_pdf(df, OUTPUT_PDF)
        make_judge_pdf(df, JUDGE_OUTPUT_PDF)
        make_register_pdf(df, REGISTER_OUTPUT_PDF)
