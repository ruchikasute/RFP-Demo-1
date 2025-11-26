import re
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import os, io
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

def _create_paragraph_after(existing_paragraph, text=None, apply_default_spacing=True):
    new_p_elm = OxmlElement("w:p")
    existing_paragraph._p.addnext(new_p_elm)
    new_para = Paragraph(new_p_elm, existing_paragraph._parent)

    if text:
        run = new_para.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(11)

    pf = new_para.paragraph_format

    if apply_default_spacing:
        pf.space_before = Pt(4)
        pf.space_after = Pt(6)
    else:
        # For headings: let the Word template decide spacing
        pf.space_before = None
        pf.space_after = None

    return new_para


# import re
def style_has_numbering(doc, style_name):
    try:
        style = doc.styles[style_name]   # <-- FIX HERE
    except KeyError:
        return False

    try:
        return "numPr" in style._element.xml
    except:
        return False


def extract_block(tag, text):
    pattern = rf"<{tag}>(.*?)(?=<[A-Z][A-Z_ ]+>|$)"

    match = re.search(pattern, text, flags=re.DOTALL)
    if not match:
        return ""
    return match.group(1).strip()


def insert_formatted_text(doc, placeholder, text, resource_table_markdown_text=None):
    """
    Inserts formatted content into the Word doc based on RAW line patterns.
    Supports:
    - Markdown Headings (#, ##, ###)
    - Bullets (-, •, *)
    - Nested bullets (2+ spaces)
    - Markdown tables
    - Bold markers (**text**)
    """

    # ---- FIND PLACEHOLDER ----
    target_p = None
    for p in doc.paragraphs:
        if placeholder in p.text:
            target_p = p
            break

    if not target_p:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            target_p = p
                            break

    if not target_p:
        return

    target_p.text = ""
    last_para = target_p
    lines = text.splitlines()

    # Detect whether List Bullet styles have numbering
    list_bullet_has_num = style_has_numbering(doc, "List Bullet")
    list_bullet2_has_num = style_has_numbering(doc, "List Bullet 2")



    i = 0

    # ----------------------------------------
    # MAIN LOOP
    # ----------------------------------------
    while i < len(lines):

        raw = lines[i]
        stripped = raw.strip()

        if not stripped:
            i += 1
            continue

        # ----------------------------------------------------
        # 1. DETECT MARKDOWN TABLE
        # ----------------------------------------------------
        if "|" in stripped and stripped.count("|") >= 2:
            table_lines = [stripped]
            j = i + 1

            while j < len(lines) and "|" in lines[j]:
                table_lines.append(lines[j].strip())
                j += 1

            insert_markdown_table_after(doc, last_para, "\n".join(table_lines))
            i = j
            continue

        # ----------------------------------------
        # 2. DETECT MARKDOWN HEADINGS
        # ----------------------------------------
        is_md_h1 = stripped.startswith("# ")
        is_md_h2 = stripped.startswith("## ")
        is_md_h3 = stripped.startswith("### ")

        if is_md_h3:
            clean = stripped.replace("### ", "").strip()
            para = _create_paragraph_after(last_para, clean)
            
            para.style = "Heading 3"
            last_para = para
            i += 1
            continue

        if is_md_h2:
            clean = stripped.replace("## ", "").strip()
            para = _create_paragraph_after(last_para, clean)


            para.style = "Heading 2"
            last_para = para
            i += 1
            continue

        if is_md_h1:
            clean = stripped.replace("# ", "").strip()
            para = _create_paragraph_after(last_para, clean)

            para.style = "Heading 1"
            last_para = para
            i += 1
            continue

        # --------------------------------------------------------
        # FORCE PARAGRAPH BREAK AFTER ANY LINE ENDING WITH ":"
        # --------------------------------------------------------
        if stripped.endswith(":"):
            # insert this line as its own paragraph
            para = _create_paragraph_after(last_para, stripped)
            para.style = "Normal"
            last_para = para

            # insert a blank line to separate bullets
            last_para = _create_paragraph_after(last_para, "")
            i += 1
            continue

        # ----------------------------------------
        # DETECT NUMBERED HEADINGS (H2 and H3)
        # ----------------------------------------
        num_heading = re.match(r"^(\d+(?:\.\d+)+)\s+(.+)$", stripped)

        if num_heading:
            numbering = num_heading.group(1)
            title_text = num_heading.group(2)

            level = numbering.count(".") + 1

            # Heading paragraph — DO NOT apply python spacing
            para = _create_paragraph_after(
                last_para, stripped, apply_default_spacing=False
            )

            if level == 2:
                para.style = "Style2"
            else:
                para.style = "Style3"

            last_para = para
            i += 1
            continue



        # ----------------------------------------
        # 3. DETECT BULLETS
        # ----------------------------------------
        bullet_match = re.match(r"^\s*([-•*])\s*(.+)$", raw)
        is_bullet = bullet_match is not None

        indent_level = len(raw) - len(raw.lstrip(" "))

        # ----------------------------------------
        # 4. CREATE PARAGRAPH WITH CORRECT STYLE
        # ----------------------------------------
        if is_bullet:

            # Insert blank when previous is not bullet
            if i > 0:
                prev = lines[i - 1].strip()
                # if prev == "" or not re.match(r"^\s*[-•*]\s*", prev):
                #     last_para = _create_paragraph_after(last_para, "")

            # Create new bullet paragraph
            para = _create_paragraph_after(last_para, "")
            
            # APPLY REAL WORD BULLET STYLE
            if indent_level < 2:
                para.style = "List Bullet 2"
            else:
                para.style = "List Bullet 2"

            # DO NOT REMOVE NUMBERING RUNS – Word needs them
            for r in list(para._p.iter("w:r")):
                txt = "".join(t.text for t in r.iter("w:t"))
                if txt.strip() != "":
                    para._p.remove(r)
            # This preserves <w:numPr> so Word bullets appear correctly

            # Extract the actual bullet text
            _, clean = bullet_match.groups()
            clean = clean.strip()

            # Bold or normal text
            if "**" in clean:
                segments = re.split(r"\*\*(.+?)\*\*", clean)
                for idx, seg in enumerate(segments):
                    run = para.add_run(seg)
                    run.font.name = "Arial"
                    run.font.size = Pt(11)
                    if idx % 2 == 1:
                        run.bold = True
            else:
                run = para.add_run(clean)
                run.font.name = "Arial"
                run.font.size = Pt(11)

        else:
            # Normal paragraph
            para = _create_paragraph_after(last_para, stripped)
            para.style = "Normal"

        last_para = para
        i += 1
        # # ----------------------------------------
        # # 3. DETECT BULLETS
        # # ----------------------------------------
        # # is_bullet = bool(re.match(r"^\s*[-•*]\s+", raw))
        # bullet_match = re.match(r"^\s*([-•*])\s*(.+)$", raw)
        # is_bullet = bullet_match is not None


        # indent_level = len(raw) - len(raw.lstrip(" "))

        # # ----------------------------------------
        # # 4. CREATE PARAGRAPH WITH CORRECT STYLE
        # # ----------------------------------------
        # if is_bullet:

        #     if i > 0:
        #         previous_raw = lines[i - 1].strip()

        #         # Case 1: previous is NOT a bullet
        #         # Case 2: previous is BLANK (critical fix)
        #         if previous_raw == "" or not re.match(r"^\s*[-•*]\s*", previous_raw):
        #             last_para = _create_paragraph_after(last_para, "")


        #     # Now create bullet paragraph
        #     # para = _create_paragraph_after(last_para, "")
        #     para = _create_paragraph_after(last_para, None)




        #     # Apply bullet style BEFORE inserting text
        #     if indent_level < 2:
        #         para.style = "List Bullet"
        #     else:
        #         para.style = "List Bullet 2"

        #     # Extract bullet text
        #     _, clean = bullet_match.groups()
        #     clean = clean.strip()




        #     # Remove default run but KEEP bullet marker
        #     # Remove only runs — keep bullet numbering (numPr)
        #     for r in list(para._p.iter("w:r")):
        #         para._p.remove(r)



        #     # --------------------------------------------------------
            # INSERT BULLET TEXT (fallback bullet + bold support)
            # --------------------------------------------------------

        #     # Determine if numbering exists in style
        #     if indent_level < 2:
        #         numbering_exists = list_bullet_has_num
        #     else:
        #         numbering_exists = list_bullet2_has_num

        #     # If numbering missing, manually prefix visual bullet
        #     final_text = clean
        #     if not numbering_exists:
        #         final_text = "• " + final_text

        #     # Insert text with bold handling
        #     if "**" in final_text:
        #         bold_pattern = r"\*\*(.+?)\*\*"
        #         segments = re.split(bold_pattern, final_text)

        #         for idx, seg in enumerate(segments):
        #             run = para.add_run(seg)
        #             run.font.name = "Arial"
        #             run.font.size = Pt(11)
        #             if idx % 2 == 1:
        #                 run.bold = True
        #     else:
        #         run = para.add_run(final_text)
        #         run.font.name = "Arial"
        #         run.font.size = Pt(11)


        # else:

        #     # Normal paragraph
        #     para = _create_paragraph_after(last_para, stripped)

        #     para.style = "Normal"

        #     # Bold handling
        #     if "**" in stripped:
        #         bold_pattern = r"\*\*(.+?)\*\*"
        #         para.clear()
        #         segments = re.split(bold_pattern, stripped)
        #         for idx, seg in enumerate(segments):
        #             run = para.add_run(seg)
        #             run.font.name = "Arial"
        #             run.font.size = Pt(11)
        #             if idx % 2 == 1:
        #                 run.bold = True

  


def insert_markdown_table_after(doc, last_para, table_text):
    """
    Converts a markdown-style table into a Word table and inserts it after last_para.
    Entire formatting comes from Word style 'Style1'.
    """

    rows = [r.strip() for r in table_text.split("\n") if r.strip()]
    if len(rows) < 2:
        return None

    # Parse header + data
    headers = [h.strip() for h in rows[0].strip("| ").split("|")]
    data_rows = [
        [c.strip() for c in r.strip("| ").split("|")]
        for r in rows[2:]
    ]

    # Create table (Word will apply Style1 styling)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Style1"   # 🔥 Apply your custom table style
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # --- HEADER ---
    hdr_cells = table.rows[0].cells
    for ci, head in enumerate(headers):
        hdr_cells[ci].text = head
        # Make header bold (Style1 may already do this)
        for run in hdr_cells[ci].paragraphs[0].runs:
            run.font.bold = True
        hdr_cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- BODY ROWS ---
    for r in data_rows:
        row_cells = table.add_row().cells
        for ci, cell_text in enumerate(r):
            if ci < len(row_cells):
                row_cells[ci].text = cell_text

    # Insert table in the document after placeholder paragraph
    last_para._p.addnext(table._tbl)

    return table

from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

def insert_image_at_placeholder(doc, placeholder, image_bytes):
    """
    Finds <PPT_IMAGE> placeholder and replaces it with an image.
    """
    target_para = None

    # 1. Search in normal paragraphs
    for p in doc.paragraphs:
        if placeholder in p.text:
            target_para = p
            break

    # 2. Search in tables if not in body
    if not target_para:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            target_para = p
                            break

    if not target_para:
        return  # placeholder not found

    # Delete placeholder text
    target_para.text = ""

    # Insert image in the same paragraph
    run = target_para.add_run()
    run.add_picture(io.BytesIO(image_bytes), width=Inches(7.5))   # adjust size

    # Optional styling
    target_para.alignment = 1   # center alignment


def insert_plain_preview(doc, placeholder, preview_text):
    """
    Inserts text exactly as shown in preview:
    - empty lines preserved
    - headings detected as plain bold text (if surrounded by **)
    - paragraphs inserted as-is
    """
    target_p = None
    for p in doc.paragraphs:
        if placeholder in p.text:
            target_p = p
            break

    if not target_p:
        return

    # clear placeholder paragraph
    target_p.text = ""
    last = target_p

    blocks = preview_text.split("\n\n")

    for block in blocks:
        lines = block.strip().split("\n")

        # If block starts with **Heading**
        if len(lines) == 1 and lines[0].startswith("**") and lines[0].endswith("**"):
            heading = lines[0][2:-2].strip()
            para = _create_paragraph_after(last, heading)
            para.style = "Heading 3"
            last = para
            continue

        # Normal paragraph
        para = _create_paragraph_after(last, block.strip())
        para.style = "Normal"
        last = para


def normalize_heading_indents(doc):
    """
    Fixes huge indentation by resetting Heading 2 and Heading 3 paragraph indents.
    """
    for heading_style in ["Heading 2", "Heading 3"]:
        style = doc.styles[heading_style]
        pf = style.paragraph_format
        
        pf.left_indent = Pt(0)            # No left indent
        pf.first_line_indent = Pt(0)      # No first-line indent
        pf.space_before = Pt(10)
        pf.space_after = Pt(6)
        # YOU CAN TWEAK THIS FOR YOUR TEMPLATE

from docx.shared import Pt, RGBColor

def update_heading2_style(doc):
    """
    Overrides the Heading 2 style formatting while keeping it as Heading 2.
    """

    style = doc.styles["Heading 2"]
    font = style.font

    # Customize here:
    font.name = "Arial"
    font.size = Pt(14)                 # Font size
    font.bold = True                   # Optional
    font.color.rgb = RGBColor(0, 32, 96)   # Dark blue numbers, adjust as needed

    # Adjust indentation
    p_format = style.paragraph_format
    p_format.left_indent = Pt(0)       # remove huge indent
    p_format.first_line_indent = Pt(0)
    p_format.space_before = Pt(15)
    p_format.space_after = Pt(4)

