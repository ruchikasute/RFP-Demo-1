import streamlit as st
import pandas as pd
import re
import glob
import io
import os
from pptx import Presentation
from docx import Document
from openai import AzureOpenAI
from dotenv import load_dotenv
import re



# --- Load your .env file safely ---
load_dotenv()

# # --- Normalize environment variables for Azure SDK ---
# # These three lines make sure AzureOpenAI gets what it expects
# os.environ["AZURE_OPENAI_API_KEY"] = os.getenv("AZURE_OPENAI_FRFP_KEY")
# os.environ["AZURE_OPENAI_ENDPOINT"] = "https://craveopenai.openai.azure.com/"
# os.environ["AZURE_OPENAI_API_VERSION"] = os.getenv("AZURE_OPENAI_FRFP_VERSION")

# # --- Initialize the Azure OpenAI client ---
# try:
#     client = AzureOpenAI(
#         azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
#         api_key=os.environ["AZURE_OPENAI_API_KEY"],
#         api_version=os.environ["AZURE_OPENAI_API_VERSION"]
#     )
#     # st.info("✅ Connected to Azure OpenAI successfully.")
# except Exception as e:
#     st.error(f"⚠️ Azure OpenAI connection failed: {e}")


# ============================================================
# Helper Functions
# ============================================================

def call_llm(prompt, client, model_name):
    """Call Azure OpenAI model."""
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"



def extract_ppt_text(ppt_path):
    """
    Extract readable text from PPT (grouped shapes + tables) and detect both
    'Working Together' slides (Objects & ABAP Programs).
    """
    import re
    text = ""
    # working_together_objects = ""
    # working_together_abap = ""
    prs = Presentation(ppt_path)

    def extract_from_shape(shape):
        content = ""
        if hasattr(shape, "text") and shape.text.strip():
            content += shape.text.strip() + "\n"
        if hasattr(shape, "shapes"):  # recurse into grouped shapes
            for sub_shape in shape.shapes:
                content += extract_from_shape(sub_shape)
        if shape.shape_type == 19:  # handle tables
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        content += cell.text.strip() + "\n"
        return content

    for slide in prs.slides:
        slide_text = ""
        for shape in slide.shapes:
            slide_text += extract_from_shape(shape)
        clean_text = slide_text.strip()
        if not clean_text:
            continue

        # if re.search(r"working\s*together", clean_text, re.IGNORECASE):
        #     if "objects" in clean_text.lower():
        #         working_together_objects = clean_text
        #     elif re.search(r"abap\s*program", clean_text, re.IGNORECASE):
        #         working_together_abap = clean_text

        text += clean_text + "\n\n"

    # if working_together_objects or working_together_abap:
    #     st.success("✅ 'Working Together' slides successfully extracted from PPT.")
    # else:
    #     st.warning("⚠️ Could not detect any 'Working Together' slides — check slide text formatting.")

    # return text.strip(), working_together_objects.strip(), working_together_abap.strip()
    return text.strip()


from docx.shared import Pt

def insert_text(doc, heading_title, text_block):
    """Adds a page break, heading, and paragraphs from a text block."""
    if not text_block:
        return  # skip if empty
    doc.add_page_break()
    doc.add_heading(heading_title, level=1)
    for line in text_block.split("\n"):
        clean = line.strip()
        if not clean:
            continue
        p = doc.add_paragraph(clean)
        p_format = p.paragraph_format
        p_format.space_after = Pt(6)
        p_format.line_spacing = 1.2
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)



def insert_annexure_table(doc, placeholder, df):
    """Insert an Annexure-style table (Object, Issue, Key Modernization Steps) into the placeholder."""
    inserted = False

    for para in doc.paragraphs:
        if placeholder in para.text:
            inserted = True
            para.text = ""  # clear placeholder text

            # --- Create table ---
            table = doc.add_table(rows=1, cols=3)
            table.style = "Table Grid"

            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "Object Name"
            hdr_cells[1].text = "Issue"
            hdr_cells[2].text = "Key Modernization Steps"

            # --- Format header cells ---
            for cell in hdr_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                cell.width = Pt(200)

            # --- Populate rows from DataFrame ---
            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row.get("object name", row.get("Object Name", "")))
                row_cells[1].text = re.sub(r"<[^>]+>", "", str(row.get("issue", row.get("Issue", ""))))
                row_cells[2].text = re.sub(r"<[^>]+>", "", str(row.get("key modernization steps", row.get("Key Modernization Steps", ""))))

            para._element.addnext(table._element)
            break

    if not inserted:
        st.warning("⚠️ No <<ANNEXURE>> placeholder found. Appending Annexure at the end.")
        doc.add_page_break()
        doc.add_heading("Annexure — Modernization Object Summary", level=1)
        table = doc.add_table(rows=1, cols=3)
        table.style = "Table Grid"

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Object Name"
        hdr_cells[1].text = "Issue"
        hdr_cells[2].text = "Key Modernization Steps"

        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row.get("object name", row.get("Object Name", "")))
            row_cells[1].text = re.sub(r"<[^>]+>", "", str(row.get("issue", row.get("Issue", ""))))
            row_cells[2].text = re.sub(r"<[^>]+>", "", str(row.get("key modernization steps", row.get("Key Modernization Steps", ""))))
def add_working_together_table(doc, heading, slide_text):
    """
    Add a Working Together slide section in tabular format to the Word document.
    """
    import pandas as pd
    from docx.shared import Pt

    if not slide_text:
        return

    # Split into key/value lines
    lines = [l.strip() for l in slide_text.split("\n") if l.strip()]
    data = []
    for line in lines:
        if ":" in line:
            parts = line.split(":", 1)
            data.append((parts[0].strip(), parts[1].strip()))
        elif "–" in line:
            parts = line.split("–", 1)
            data.append((parts[0].strip(), parts[1].strip()))
        else:
            data.append(("", line))

    df = pd.DataFrame(data, columns=["Category", "Details"])

    # Add heading and table to the doc
    doc.add_page_break()
    doc.add_heading(heading, level=1)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Category"
    hdr_cells[1].text = "Details"

    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    for _, row in df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row["Category"])
        cells[1].text = str(row["Details"])

    # Format all cells
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)


from docx.shared import Inches

def insert_sustainability_section(doc, image_top=None, image_bottom=None):
    """
    Inserts a sustainability and EcoVadis section before 'Project Scope' section.
    Ensures correct order: Top image → Text → Bottom image.
    """
    from docx.shared import Inches, Pt
    import re, os

    sustainability_text = """
At Crave InfoTech, we recognize the importance of sustainability in fostering long-term growth, innovation, and responsibility. 
As a leading player in the IT sector, we are committed to integrating sustainable practices into every aspect of our business. 
Our approach to sustainability is guided by the principles of environmental stewardship, social responsibility, and economic viability.

**Ecovadis Rating for Crave InfoTech**
Crave InfoTech maintains a strong commitment to sustainability excellence, as reflected in our EcoVadis rating. 
This recognition underscores our focus on ethical business practices, sustainable procurement, and corporate responsibility across all engagements.
    """

    # 🔍 Find insertion point: before "Project Scope"
    target_para = None
    for para in doc.paragraphs:
        if re.search(r"\bProject\s+Scope\b", para.text, re.IGNORECASE):
            target_para = para
            break

    if target_para:
        parent = target_para._element.getparent()
        idx = parent.index(target_para._element)

        # --- Create paragraph for top image, if exists ---
        if image_top and os.path.exists(image_top):
            p_top = doc.add_paragraph()
            run = p_top.add_run()
            run.add_picture(image_top, width=Inches(6))
            parent.insert(idx, p_top._element)
            idx += 1  # move index forward

        # --- Sustainability text ---
        p_text = doc.add_paragraph(sustainability_text.strip())
        p_text.paragraph_format.space_after = Pt(6)
        for run in p_text.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)
        parent.insert(idx, p_text._element)
        idx += 1

        # --- Bottom image, if exists ---
        if image_bottom and os.path.exists(image_bottom):
            p_bottom = doc.add_paragraph()
            run = p_bottom.add_run()
            run.add_picture(image_bottom, width=Inches(3))
            parent.insert(idx, p_bottom._element)

        st.success("🌱 Sustainability section (with EcoVadis images) inserted before 'Project Scope'.")
    else:
        st.warning("⚠️ 'Project Scope' not found — appending sustainability section at document end.")
        doc.add_page_break()
        doc.add_paragraph(sustainability_text.strip())


def add_coreassess_pricing_tables(doc):
    """
    Inserts the CoreAssess pricing tables right after the 'Commercials' section heading.
    If not found, appends them at the end.
    """
    from docx.shared import RGBColor, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tiers = [
        ("Starter Pack", "Assess 5 ABAP Objects", "1 week", "Complimentary"),
        ("Silver", "Assess 50+ ABAP Objects", "1–2 weeks", "$10 per Object"),
        ("Gold", "Assess 2500+ ABAP Objects", "2–4 weeks", "$7.5 per Object"),
        ("Platinum", "Assess 10000+ ABAP Objects", "4+ weeks", "$5 per Object"),
    ]

    # --- Find 'Commercials' heading ---
    target_para = None
    for para in doc.paragraphs:
        if re.search(r"\bCommercials\b", para.text, re.IGNORECASE):
            target_para = para
            break

    # --- Create the pricing table ---
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    headers = ["Tier", "Scope", "Duration", "Price"]

    # Header styling
    for i, h in enumerate(headers):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
        tc_pr = hdr[i]._element.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "0072C6")  # Blue header
        tc_pr.append(shd)

    # Data rows
    for title, scope, duration, price in tiers:
        row_cells = table.add_row().cells
        row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text = (
            title,
            scope,
            duration,
            price,
        )
        for c in row_cells:
            tc_pr = c._element.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "E7EEF7")  # Light gray rows
            tc_pr.append(shd)
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(3)
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(10)

    # --- Insert the table right after 'Commercials' ---
    if target_para:
        parent = target_para._element.getparent()
        idx = parent.index(target_para._element)
        parent.insert(idx + 1, table._element)
        st.info("📊 Pricing table inserted below 'Commercials' section.")
    else:
        st.warning("⚠️ 'Commercials' section not found. Appending pricing table at the end.")
        doc.add_page_break()
        doc.add_heading("Working Together — ABAP Objects", level=1)
        doc._body._body.append(table._element)


    # Repeat for ABAP Programs
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def insert_formatted_text(doc, placeholder, raw_text):
    """
    Smart parser for CoreAssess SOW: handles any LLM output format 
    (Markdown, numbered, or plain text) and converts it to rich Word formatting.
    """
    def set_cell_shading(cell, fill_color):
        """Add shading to a table cell."""
        tc_pr = cell._element.tcPr
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill_color)
        tc_pr.append(shd)

    # def insert_styled_table(parent, headers, rows):
    #     """Insert styled table with blue header and light gray body."""
    #     table = parent.add_table(rows=len(rows) + 1, cols=len(headers))
    #     table.style = "Table Grid"
    #     hdr_cells = table.rows[0].cells
    #     for i, h in enumerate(headers):
    #         hdr_cells[i].text = h.strip()
    #         set_cell_shading(hdr_cells[i], "0072C6")
    #         for run in hdr_cells[i].paragraphs[0].runs:
    #             run.font.bold = True
    #             run.font.color.rgb = RGBColor(255, 255, 255)
    #     for r, row_data in enumerate(rows):
    #         cells = table.rows[r + 1].cells
    #         for c, val in enumerate(row_data):
    #             cells[c].text = str(val).strip()
    #             set_cell_shading(cells[c], "E7EEF7")
    #     return table
    def insert_styled_table(parent, headers, rows):
        """Insert styled table with blue header and light gray body."""
        from docx.shared import RGBColor

        table = parent.add_table(rows=len(rows) + 1, cols=len(headers))
        table.style = "Table Grid"

        # --- Header Row ---
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h.strip()
            set_cell_shading(hdr_cells[i], "0072C6")
            for run in hdr_cells[i].paragraphs[0].runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

        # --- Body Rows ---
        for r, row_data in enumerate(rows):
            cells = table.rows[r + 1].cells
            for c in range(len(headers)):
                try:
                    val = row_data[c]
                except IndexError:
                    val = ""  # <-- Fix: fill blank for missing columns
                cells[c].text = str(val).strip()
                set_cell_shading(cells[c], "E7EEF7")

        return table


    def add_heading(doc, text, level):
        """Add heading with correct style and spacing."""
        p = doc.add_paragraph(text.strip(), style=f"Heading {level}")
        p.paragraph_format.space_after = Pt(4)
        return p

    inserted = False
    for para in doc.paragraphs:
        if placeholder in para.text:
            inserted = True
            parent = para._element.getparent()
            idx = parent.index(para._element)
            parent.remove(para._element)

            lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
            new_elements = []
            i = 0
            while i < len(lines):
                line = lines[i]

                # === Numbered Section Headings ===
                if re.match(r"^\d+\.\s+[A-Z]", line):
                    heading_text = re.sub(r"^\d+\.\s*", "", line)
                    new_elements.append(add_heading(doc, heading_text, 1)._element)
                    i += 1
                    continue

                # === Subheadings (e.g., 5.1 or bold headings) ===
                if re.match(r"^\d+\.\d+\s+[A-Z]", line):
                    heading_text = re.sub(r"^\d+\.\d+\s*", "", line)
                    new_elements.append(add_heading(doc, heading_text, 2)._element)
                    i += 1
                    continue

                # === Markdown / Pipe Tables ===
                if line.startswith("|") and "|" in line:
                    table_lines = []
                    while i < len(lines) and lines[i].startswith("|"):
                        table_lines.append(lines[i])
                        i += 1
                    headers = [h.strip("* ") for h in table_lines[0].strip("|").split("|")]
                    rows = [r.strip("|").split("|") for r in table_lines[2:]]
                    table = insert_styled_table(doc, headers, rows)
                    new_elements.append(table._element)
                    continue

                # === Bullets ===
                if line.startswith("- ") or line.startswith("• "):
                    bullet_text = re.sub(r"^[-•]\s*", "", line)
                    p = doc.add_paragraph(bullet_text, style="List Bullet 2")
                    p.paragraph_format.left_indent = Pt(18)
                    p.paragraph_format.space_after = Pt(2)
                    new_elements.append(p._element)
                    i += 1
                    continue

                # === Regular Paragraph ===
                p = doc.add_paragraph(line)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.2
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(11)
                new_elements.append(p._element)
                i += 1

            # insert parsed content before next element
            for el in reversed(new_elements):
                parent.insert(idx, el)
            break

    if not inserted:
        st.warning(f"⚠️ Placeholder {placeholder} not found — appending content at end.")
        doc.add_paragraph(raw_text)

# ============================================================
# Core Function
# ============================================================

def generate_sow(df, client, model_name, client_name=None, repo_dir="Knowledge_Repo/Coreassess_KR"):
    """Generate full SOW docx directly."""
    client_ref = client_name if client_name else "the Client"

    # Find available PPT references
    ppt_files = glob.glob(os.path.join(repo_dir, "*.pptx"))
    if not ppt_files:
        ppt_text = "No PPTs found."
        chosen_ppt = "None"
    else:
        # ppt_text, working_together_objects, working_together_abap = extract_ppt_text(ppt_files[0])
        ppt_text = extract_ppt_text(ppt_files[0])
        chosen_ppt = ppt_files[0]


    # Build prompt
    total = len(df)

    # Safe extraction for sample issues
    if "issue" in df.columns:
        sample_col = df["issue"]
    else:
        sample_col = df.iloc[:, 0]

    sample_issues = "; ".join(sample_col.astype(str).tolist()[:5])


    prompt = f"""
You are a Senior SAP consultant from Crave InfoTech preparing a professional
Statement of Work (SOW) for a Clean Core Assessment (CoreAssess.AI) engagement with {client_ref}.

Below is Crave Infotech’s internal knowledge reference extracted from our Clean Core Assessment repository (PPT/Knowledge Base). 
Use it to infer tone, structure, and technical depth:

REFERENCE MATERIAL:
{ppt_text}

Using this reference as context, generate a **comprehensive, client-ready Statement of Work** structured as follows:

1. Introduction  
   - Provide a high-level narrative overview of the Clean Core Assessment initiative, its purpose, and value to {client_ref}.  
   - Explain how Crave Infotech’s CoreAssess.AI helps organizations modernize ABAP custom objects, identify technical debt, and align with SAP’s Clean Core strategy.  
   - Mention Crave’s SAP Build Process Automation (SBPA), AI-driven analysis, and clean-core accelerators that enhance assessment accuracy, performance, and compliance.  
   - Use a formal, consultative tone that emphasizes Crave Infotech’s delivery capability and alignment with SAP’s modernization roadmap.  

2. Project Scope  
   2.1 In-Scope Items  
   - Clearly outline activities under the Clean Core Assessment scope — e.g., object evaluation, extensibility classification (On-Stack, Side-by-Side, Retire), and modernization recommendations.  
   2.2 Solution Architecture  
   - Describe the technical architecture and toolset used — including CoreAssess.AI platform, ABAP parser, and integration with BTP or Solution Manager.  
   2.3 Prerequisites and Key Assumptions  
   - List all assumptions (e.g., system access, readiness of transport data, and availability of ABAP repository).  
   2.4 Out of Scope  
   - Clearly state excluded items (e.g., remediation implementation, non-ABAP system assessments).  
   2.5 Deliverables  
   - Summarize deliverables such as Assessment Report, Insights Summary, Recommendation Deck, and Modernization Plan.  

3. Key Insights and Recommendations  
   - Using the provided data summary below, identify key patterns in ABAP object issues and modernization approaches.  
   - Provide categorized recommendations:  
       3.1. On-Stack Extensibility  
       3.2. Side-by-Side Extensibility  
       3.3. Retire Candidates  
   - Reference technical metrics and business rationale, focusing on cost optimization and compliance benefits.  
   - End this section with a summary paragraph linking findings to SAP’s Clean Core strategy.  

   Total Objects: {total}  
   Example Issues: {sample_issues}  

4. Benefits of CoreAssess.AI  
   - Compare Crave’s AI-driven assessment with traditional manual clean-core evaluations.  
   - Highlight benefits like automated code scanning, structured modernization mapping, ROI analysis, and faster turnaround time.  
   - Emphasize measurable business outcomes for {client_ref} — improved performance, reduced technical debt, and audit-ready modernization planning.  

5. Project Delivery Approach  
   5.1 Project Organization Structure  
   - Describe Crave’s typical project governance and communication model for a Clean Core Assessment engagement.  
   5.2 Project Resource Planning  
   - Add a short description of resource allocation and Crave’s hybrid delivery model (onsite–offshore mix).  
   - Then include a **Project Resource Planning Table** with columns:  
        *Project Phase*, *Functional Consultant*, *Technical Consultant*, *Other Roles (Developer, Tester, PM, Architect, etc.)*  
     and rows such as:  
        - Project Preparation  
        - Discovery & Object Extraction  
        - Assessment & Categorization  
        - Insights & Recommendation  
        - Report Preparation & Review  
        - Presentation & Sign-off  
   5.3 Implementation Methodology  
   - Outline the phased approach Crave follows, from assessment kickoff to presentation and handover.  
   5.4 Communication Plan  
   - Describe how Crave and {client_ref} will communicate and manage progress throughout the project.  
   - Include these tables:  
       **Communication Schedule** (Interaction | Frequency | Purpose)  
       **Issue Management and Escalation** (Task | Timescale | Responsibility)  
       **Issue Classification** (Severity | Definition | Reporting | Resolution Owner)  
       **Escalation Process** (Issue Type | Escalation Point | Criteria | Governance Role)  

6. Timelines  
   6.1 Delivery Timeliness  
   - Provide an indicative duration for each phase (typically 3–6 weeks total).  
   6.2 Efforts and Resource Allocation  
   - Present a short summary of resource utilization (Consultants, PM, ABAP Specialist, QA).  
   6.3 Commercials  
   - Describe how Crave offers flexible engagement options (e.g., per-object, per-phase, or fixed-scope pricing).  
   6.4 Payment Terms  
   - List typical payment milestones (e.g., Kickoff – 20%, Report Delivery – 50%, Final Presentation – 30%) in tabular format.  

7. Sign-Off  
   - Add formal sign-off and acceptance text for both Crave Infotech and {client_ref}.  
   - *Party*, *Designation*, *Signature*, *Date*.  

8. Other Assumptions  
   8.1 Dependencies  
   - List client-side dependencies (e.g., system access, object data extraction support).  
   8.2 Limitations  
   - Mention constraints (e.g., tool version, scope limits, data quality).  
   8.3 General Provisions  
   - Add closing statements about intellectual property, confidentiality, and engagement validity.

---

Formatting Instructions:
- Use numbered section headings exactly as shown above.  
- Each section must contain well-written paragraphs — avoid short bullets except inside structured tables.  
- Maintain Crave Infotech’s corporate tone: **formal, confident, and consultative**.  
- Avoid generic wording or references to other organizations (e.g., “Oatey Co.”).  
- Begin the document directly with the section heading **“1. Introduction”** — do not include titles like “Statement of Work” or “Proposal for …”.  
- Total length: around **5–6 Word pages**.
"""


    # Get LLM result
    # --- Split SOW by numbered headings like "1. Executive Summary" ---

    full_sow = call_llm(prompt, client, model_name)


    # --- Use Template ---
    template_path = "Template/CoreAssess_Template.docx"

    if os.path.exists(template_path):
        doc = Document(template_path)
        st.info("📄 Using Word template for SOW.")
    else:
        st.warning("⚠️ Template not found. Creating a blank document.")
        doc = Document()
         


    def insert_full_sow(doc, placeholder, sow_text):
        """Insert AI-generated SOW content right before hardcoded sections."""
        from docx import Document as NewDoc
        from docx.shared import Pt

        inserted = False
        for para in doc.paragraphs:
            if placeholder in para.text:
                inserted = True
                para.text = ""

                # Create a temporary doc for new content
                temp_doc = NewDoc()
                lines = [line.strip() for line in sow_text.split("\n") if line.strip()]

                for line in lines:
                    # Headings like "1. Executive Summary"
                    if re.match(r"^\s*\d+\.\s+[A-Z]", line):
                        heading_text = re.sub(r"^\s*\d+\.\s*", "", line).strip()
                        temp_doc.add_page_break()
                        temp_doc.add_heading(heading_text, level=1)
                    # Bullet points
                    elif line.startswith("- ") or line.startswith("• "):
                        bullet_text = line[2:].strip() if line.startswith("- ") else line[1:].strip()
                        p = temp_doc.add_paragraph(f"• {bullet_text}")
                        p.paragraph_format.left_indent = Pt(18)
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.line_spacing = 1.15
                    # Normal paragraph
                    else:
                        p = temp_doc.add_paragraph(line)
                        p_format = p.paragraph_format
                        p_format.space_after = Pt(6)
                        p_format.line_spacing = 1.2
                        for run in p.runs:
                            run.font.name = "Calibri"
                            run.font.size = Pt(11)

                # Insert right before the next element in Word XML
                parent = para._element.getparent()
                idx = parent.index(para._element)
                for new_para in reversed(temp_doc.paragraphs):
                    parent.insert(idx, new_para._element)

                break

        if not inserted:
            st.warning("⚠️ Placeholder <<CONTENT START>> not found. Appending at end.")
            for line in sow_text.split("\n"):
                if line.strip():
                    doc.add_paragraph(line.strip())



    # # --- Insert generated content ---
    # insert_full_sow(doc, "<<CONTENT START>>", full_sow)
    insert_formatted_text(doc, "<<CONTENT START>>", full_sow)
    # --- Add Sustainability section below Executive Summary ---



    add_coreassess_pricing_tables(doc)
    insert_sustainability_section(
        doc,
        image_top="Images/Crave Awards.png",   # optional top banner
        image_bottom="Images/Sustainability.png"       # optional EcoVadis badge
    )
    
    # --- Add Annexure section at the end ---
    doc.add_page_break()
    doc.add_heading("Annexure — Modernization Object Summary", level=1)

    # Create table with 3 columns: Object, Issue, Modernization Steps
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Object Name"
    hdr_cells[1].text = "Issue"
    hdr_cells[2].text = "Key Modernization Steps"

    # Format header
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    # Fill data from uploaded Excel (df)
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row.get("object name", row.get("Object Name", "")))
        row_cells[1].text = str(row.get("issue", row.get("Issue", "")))
        row_cells[2].text = str(row.get("key modernization steps", row.get("Key Modernization Steps", "")))

    # Add a closing note
    doc.add_paragraph()
    doc.add_paragraph("This annexure provides a detailed mapping of identified objects, their issues, and the corresponding modernization steps proposed by CoreAssess.AI.")


    # --- Save to memory ---
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    # --- Preview + Download ---
    # st.markdown("### 📄 Preview of Generated SOW")
    # preview_text = "\n".join(full_sow.split("\n")[:50])
    # st.text(preview_text.strip())

    st.success(f"✅ SOW generated using `{os.path.basename(chosen_ppt)}` and inserted into template.")
    st.download_button(
        label="📥 Download SOW Document (.docx)",
        data=buffer,
        file_name=f"{client_ref.replace(' ', '_')}Coreassess_SOW_.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ============================================================
# Streamlit UI
# ============================================================

def main():
    st.title("🌐 CoreAssess.AI — Auto SOW Generator")

    client_name = st.text_input("Client Name", placeholder="e.g., Adani Group")
    uploaded = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"])

    if uploaded:
        df = pd.read_excel(uploaded)
        st.success(f"✅ File `{uploaded.name}` loaded successfully!")
        st.dataframe(df.head(5))

        # Azure OpenAI setup
        client = AzureOpenAI(
            azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
            api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
            api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
        )
        model_name = "codetest"

        if st.button("⚡ Generate SOW Document"):
            generate_sow(df, client, model_name, client_name)
