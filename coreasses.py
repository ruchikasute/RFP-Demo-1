import streamlit as st
from openai import AzureOpenAI
import os, io
from docx import Document
from PyPDF2 import PdfReader
from datetime import datetime
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from dotenv import load_dotenv
load_dotenv()




def insert_excel_appendix(final_doc):
    """
    Inserts the uploaded Excel content (stored as markdown)
    into the <APPENDIX> placeholder inside the Word document.
    """
    if "excel_markdown" not in st.session_state:
        return  # nothing to insert

    excel_md = st.session_state["excel_markdown"]

    # IMPORTANT: Ensure placeholder exists in template
    placeholder = "<APPENDIX>"

    insert_formatted_text(
        doc=final_doc,
        placeholder=placeholder,
        text=excel_md
    )



def generate_all_sections(client, model_name,  client_name):

    prompt = f"""
You are a Senior SAP Consultant from Crave InfoTech.

Generate ALL 8 sections of an CoreAssess.AI SOW.

Client: {client_name}

REFERENCE SOW STYLE GUIDE:
{st.session_state['knowledge_text']}


STRICT RULES:
- Output ONLY the 8 sections.
- Follow EXACT tag structure.
- No extra text outside tags.

=====================================
OUTPUT FORMAT (FOLLOW EXACTLY)
=====================================

<EXEC_SUMMARY>
[content]

<ABOUT_CRAVE>
[content]

<ABOUT_CLIENT>
[content]

<PROJECT_SCOPE>
[content]

<DELIVERY_APPROACH>
[content]

<RESOURCE_TIMELINE>
Write a SINGLE short narrative paragraph (3–5 lines).

<SIGN_OFF>
[content]

<APPENDIX>
[content]

"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )

    full_output = response.choices[0].message.content.strip()

    # EXTRACT each block
    extracted = {
        "Executive Summary": extract_block("EXEC_SUMMARY", full_output),
        "About Crave InfoTech": extract_block("ABOUT_CRAVE", full_output),
        "About Client": extract_block("ABOUT_CLIENT", full_output),
        "Project Scope": extract_block("PROJECT_SCOPE", full_output),
        "Project Delivery Approach": extract_block("DELIVERY_APPROACH", full_output),
        "Resource Allocation & Timelines": extract_block("RESOURCE_TIMELINE", full_output),
        "Sign Off": extract_block("SIGN_OFF", full_output),
        "Appendix": extract_block("APPENDIX", full_output),
    }

    return extracted


import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document

# Import shared modules
from Modules.extractors import extract_text_from_file, summarize_large_rfp, extract_block
from Modules.word_insert import insert_formatted_text, insert_image_at_placeholder
from Modules.placeholders import (
    replace_client_name_in_doc,
    replace_submission_date,
    replace_inline_placeholder,
    insert_document_number,
    generate_document_number,
)
from Modules.knowledge import load_knowledge_text
from Modules.llm import regenerate_section_llm
from Modules.preview import section_preview_tabs



if "knowledge_text" not in st.session_state:
    st.session_state["knowledge_text"] = load_knowledge_text()

import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from pptx import Presentation

def generate_exec_summary(client, model_name, client_name):
    prompt = f"""
You are a Senior SAP Consultant from Crave InfoTech.

Generate ONLY the Executive Summary for a CoreAssess.AI-based Clean Core Assessment SOW.

Client: {client_name}

INSTRUCTIONS:
- Write 2–3 short paragraphs (5–6 lines each).
- Describe purpose of Clean Core assessment.
- Mention modernizing custom code, reducing technical debt, and increasing S/4 readiness.
- High-level business tone, no technical jargon.
- No markdown. No bullets.

Output ONLY the executive summary content.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_about_crave(client, model_name, client_name):
    prompt = f"""
Write the 'About Crave InfoTech' section for a Clean Core Assessment SOW.

INSTRUCTIONS:
- 2–3 lines only.
- Highlight SAP expertise, Clean Core accelerators, CoreAssess.AI capability.
- No markdown, no bullets.

Output only the content.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content.strip()

def generate_our_understanding(client, model_name, client_name):
    prompt = f"""
Generate the 'Our Understanding' section for a Clean Core Assessment SOW.

INSTRUCTIONS:
- Write 2–3 paragraphs.
- Describe understanding of customer's ERP landscape, custom code, enhancements, integrations.
- Explain need for modernization and clean-core alignment.
- No technical jargon. No bullets.

Output only the section content.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_project_scope(client, model_name, client_name):
    prompt = f"""
Write the 'Project Scope' section for a Clean Core Assessment SOW.

Include ONLY the following subsections:

4.1 Definition & Terminologies
• CoreAssess.AI – AI-driven assessment platform by Crave.
• ABAP Object – Any ABAP program/function/class analyzed.
• Assessment – Automated evaluation for complexity & clean core.
• Clean Core Compliance – Alignment with SAP clean-core strategy.
• BTP BOM – Recommended SAP BTP services for modernization.
• Modernization Recommendation – AI-generated remediation guidance.

4.2 Proposed Complimentary POC Scope
Write 5–6 lines describing platform access, onboarding, assessment purpose, and expected outcomes.

4.3 In-Scope Items
• Provide secure access to CoreAssess.AI.
• Onboarding and enablement session.
• Self-service ABAP assessment.
• Dashboard & report-based analysis.
• Review sessions on findings.

4.4 Prerequisites and Key Assumptions
• Access to SAP system & ABAP repo.
• Transport/object metadata available.
• Collaboration with NOVA’s technical team.

4.5 Out of Scope
• Implementation of recommendations.
• Assessment of non-ABAP systems.
• Conversion of more than 3 objects to RAP/CAPM.

4.6 Deliverables
• Platform access & credentials.
• Onboarding/training.
• AI-generated assessment reports.
• Summary assessment findings.
• Support during POC.

RULES:
- No markdown.
- Use only bullets “•”.
- Keep it high-level and business friendly.
- Output only these subsections.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()

def generate_solution_approach(client, model_name, client_name):
    prompt = f"""
Generate the 'Proposed Solution Approach' for a Clean Core Assessment SOW.

INSTRUCTIONS:
Write 2–3 paragraphs covering:
- CoreAssess.AI pipeline (extraction → analysis → recommendations)
- How ABAP custom code, SQL, and functional objects are evaluated
- How outputs like BOM, Functional Specs, Technical Specs, Effort Estimation are generated

NO technical jargon from PI/PO or Integration Suite.
NO markdown.

Output only the narrative content.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_team_structure(client, model_name, client_name):
    prompt = f"""
Write the 'Team Structure' section for a Clean Core Assessment SOW.

INSTRUCTIONS:
- Describe Crave team roles (PM, Functional Consultant, Technical Consultant, Architect).
- Keep it high-level.
- No tables, no markdown.

Output a 4–6 line paragraph.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content.strip()



def generate_commercials(client, model_name, client_name):
    prompt = f"""
Write ONLY the 'Commercials & Payment Terms' section.

INSTRUCTIONS:
- Keep it generic (do not put specific prices).
- Payment Terms Table:
Row 1: Milestone | Description | Payment %
Row 2: Kickoff | Signing of SOW | 20%
Row 3: Mid-Assessment | Submission of interim findings | 40%
Row 4: Final Report | Delivery of final Clean Core Assessment report | 40%

- No markdown.

Output the final content only.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content.strip()

def generate_sign_off(client, model_name, client_name):
    prompt = f"""
Generate the 'Sign Off' section for the CoreAssess.AI SOW.

INSTRUCTIONS:
- Formal acceptance language.
- Both parties acknowledge responsibilities.
- 2–3 sentences only.
- No markdown.

Output only the content.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content.strip()

def generate_key_assumptions(client, model_name, client_name):
    prompt = f"""
Write the 'Key Assumptions' section for a Clean Core Assessment SOW.

INSTRUCTIONS:
Use EXACT structure:
10.1 Dependency on Client
• 3–4 bullets

10.2 Other Assumptions
• 5–7 bullets

RULES:
- No markdown.
- Use bullet (•) only.

Output only the structured content.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_selected_sections(client, model_name, client_name, selected_sections):
    """
    Generate ONLY the sections the user selected (CoreAssess SOW).
    Uses parallel execution for speed.
    """

    # Map UI Section Names → Generation Functions
    section_generators = {
        "Executive Summary": lambda: generate_exec_summary(client, model_name, client_name),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, client_name),
        "Our Understanding": lambda: generate_our_understanding(client, model_name, client_name),
        "Project Scope": lambda: generate_project_scope(client, model_name, client_name),
        "Proposed Solution Approach": lambda: generate_solution_approach(client, model_name, client_name),
        "Team Structure": lambda: generate_team_structure(client, model_name, client_name),
        "Commercials & Payment Terms": lambda: generate_commercials(client, model_name, client_name),
        "Sign Off": lambda: generate_sign_off(client, model_name, client_name),
        "Key Assumptions": lambda: generate_key_assumptions(client, model_name, client_name)    }

    generated = {}

    # Run selected sections in parallel
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(section_generators[name]): name
            for name in selected_sections
            if name in section_generators
        }

        for future in concurrent.futures.as_completed(futures):
            sec_name = futures[future]
            try:
                content = future.result() or ""
            except Exception as e:
                content = f"(Error generating {sec_name}: {e})"

            # Clean HTML/Markdown tags if any
            content = re.sub(r"</?[^>]+>", "", content).strip()

            generated[sec_name] = content

    return generated


def main():
    st.title("🌐 CoreAssess.AI — SOW Generator")

    # Initialize vector DB only first time
    if "vector_db_ready" not in st.session_state:
        from Modules.knowledge import initialize_vector_db
        initialize_vector_db()
        st.session_state["vector_db_ready"] = True

    
    # Initialize keys early (safe even if they exist)
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)

    # -------------------------------
    # Client name input + RFP upload
    # -------------------------------
    client_name = st.text_input("Enter Client Name (required)", "")
    uploaded = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"])
    if uploaded:
  
        df = pd.read_excel(uploaded).fillna("")

         # store table in markdown so preview + Word both can use
        excel_markdown = df.to_markdown(index=False)
        st.session_state["excel_markdown"] = excel_markdown

        # st.session_state["reference_text"] = df.to_csv(index=False)


    # Azure LLM client
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "gpt-4o-mini"

    # Store globally (MUST be before using them)
    st.session_state["llm_client"] = client
    st.session_state["llm_model"] = model_name

    # ========================================
    # SECTION SELECTION (CoreAssess SOW)
    # ========================================

    st.markdown("---")
    st.subheader("📋 Select Sections to Generate")

    col1, col2 = st.columns(2)

    with col1:
        exec_summary = st.checkbox("1. Executive Summary", value=True, key="chk_exec")
        about_crave = st.checkbox("2. About Crave InfoTech", value=True, key="chk_crave")
        our_understanding = st.checkbox("3. Our Understanding", value=True, key="chk_understanding")
        project_scope = st.checkbox("4. Project Scope", value=True, key="chk_scope")

    with col2:
        solution_approach = st.checkbox("5. Proposed Solution Approach", value=True, key="chk_solution")
        team_structure = st.checkbox("6. Team Structure", value=True, key="chk_team")
        commercials = st.checkbox("8. Commercials & Payment Terms", value=True, key="chk_commercials")
        sign_off = st.checkbox("9. Sign Off", value=True, key="chk_signoff")
        assumptions = st.checkbox("10. Key Assumptions", value=True, key="chk_assumptions")

    # Build list of selected sections
    selected_sections = []
    if exec_summary: selected_sections.append("Executive Summary")
    if about_crave: selected_sections.append("About Crave InfoTech")
    if our_understanding: selected_sections.append("Our Understanding")
    if project_scope: selected_sections.append("Project Scope")
    if solution_approach: selected_sections.append("Proposed Solution Approach")
    if team_structure: selected_sections.append("Team Structure")
    if commercials: selected_sections.append("Commercials & Payment Terms")
    if sign_off: selected_sections.append("Sign Off")
    if assumptions: selected_sections.append("Key Assumptions")


    # ========================================
    # GENERATE SELECTED SECTIONS
    # ========================================

    if st.button("⚡ Generate Selected Sections"):
        st.session_state.pop("edited_sections", None)

        # Reset editors
        for key in list(st.session_state.keys()):
            if key.startswith("editor_"):
                st.session_state.pop(key)

        if not selected_sections:
            st.warning("⚠ Please select at least one section to generate.")
            st.stop()

        # Generate only selected sections
        with st.spinner(f"⏳ Generating {len(selected_sections)} section(s)..."):
            generated = generate_selected_sections(client, model_name, client_name, selected_sections)

        # ORDER FIX (always output in fixed SOW order)
        MASTER_ORDER = [
            "Executive Summary",
            "About Crave InfoTech",
            "Our Understanding",
            "Project Scope",
            "Proposed Solution Approach",
            "Team Structure",
            "Commercials & Payment Terms",
            "Sign Off",
            "Key Assumptions"
        ]


        ordered_list = []
        for section_name in MASTER_ORDER:
            if section_name in generated:
                clean = re.sub(r"</?[^>]+>", "", generated[section_name]).strip()
                ordered_list.append({"title": section_name, "content": clean})

        st.session_state["edited_sections"] = ordered_list
        st.success(f"✅ Generated {len(selected_sections)} sections!")


    # ========================================
    # PREVIEW
    # ========================================

    if "edited_sections" in st.session_state:
        section_preview_tabs()


    # ========================================
    # DOWNLOAD DOCX
    # ========================================

    if "edited_sections" in st.session_state:

        buffer = io.BytesIO()
        final_doc = Document("Template/coreassess_Template.docx")

        # Basic replacements
        replace_client_name_in_doc(final_doc, client_name)
        replace_submission_date(final_doc)
        doc_no = generate_document_number(client_name)
        insert_document_number(final_doc, "<DOCUMENT_NO>", doc_no)

        # Placeholder map
        placeholder_map = {
            "Executive Summary": "<EXEC_SUMMARY>",
            "About Crave InfoTech": "<ABOUT_CRAVE>",
            "Our Understanding": "<OUR_SOL>",
            "Project Scope": "<PROJECT_SCOPE>",
            "Proposed Solution Approach": "<DELIVERY_APPROACH>",
            "Team Structure": "<TEAM_STRUCTURE>",
            "Commercials & Payment Terms": "<PAYMENT TERMS>",
            "Sign Off": "<SIGN_OFF>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>",
            "Appendix": "<APPENDIX>"
        }


        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)

        # Excel appendix (if uploaded)
        insert_excel_appendix(final_doc)

        final_doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Download Final SOW Document",
            data=buffer,
            file_name=f"CoreAssess_SOW_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

