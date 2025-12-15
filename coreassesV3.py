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


def get_knowledge():
    """Always return latest knowledge text for LLM prompts."""
    return st.session_state.get("knowledge_text", "") or ""


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
from Modules.word_insert import insert_plain_preview, insert_markdown_table_after, insert_table_at_placeholder, remove_table_by_tag
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

def generate_solution_section(client, model_name, knowledge_text, client_name):

    placeholder = "[[ARCHITECTURE_IMG]]"

    prompt = f"""
You are creating Section 5 — Solution Overview for the Clean Core Assessment SOW for {client_name}.

MANDATORY:
You MUST include the placeholder EXACTLY as: {placeholder}
Do NOT modify it.

====================
STRUCTURE
====================

5.1 Proposed Architecture
Write a 6–8 line paragraph describing:
- High-level SAP BTP architectural positioning for Clean Core modernization
- How the platform enables extensibility, governance, and sustainable operations
- How it supports standardization and future readiness
- Include a reference to the architecture diagram (business tone only)

Then on a new line write ONLY:
{placeholder}

5.2 Bill of Material (BOM)
Write 5–7 bullets using (•).
The bullets must:
- List SAP BTP services relevant for Clean Core modernization
- Use only business-friendly names (e.g., SAP BTP Identity Authentication, SAP BTP Workflow Management)
- No descriptions after the service name
- No technical jargon
- Do NOT reuse any service more than once
- Let the LLM choose the appropriate SAP BTP services (do NOT copy from the prompt)

====================
RULES
====================
- No markdown
- No bold or formatting syntax
- Use only paragraphs + (•) bullets
- Output MUST include placeholder {placeholder} exactly once
- Output only the section content


USE THIS CONTEXT FOR UNDERSTANDING:

KNOWLEDGE BASE:
{knowledge_text}
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.35,
    )

    return response.choices[0].message.content.strip()



def generate_solution_approach(client, model_name, client_name):
    prompt = f"""
Generate the 'Project Approach' for a Clean Core Assessment SOW.

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
    # universal knowledge access
    knowledge = get_knowledge()



    # Map UI Section Names → Generation Functions
    section_generators = {
        "Executive Summary": lambda: generate_exec_summary(client, model_name, client_name),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, client_name),
        "Our Understanding": lambda: generate_our_understanding(client, model_name, client_name),
        "Project Scope": lambda: generate_project_scope(client, model_name, client_name),
        "Solution": lambda: generate_solution_section(client, model_name, knowledge, client_name),  # NEW
        "Project Approach": lambda: generate_solution_approach(client, model_name, client_name),
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

    # Azure LLM client (initialize early so generate button inside expander can use it)
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "gpt-4o-mini"

    st.session_state["llm_client"] = client
    st.session_state["llm_model"] = model_name

    # -------------------------------
    # Client + Input Configuration (expander with ordered selection)
    # -------------------------------

    with st.expander("⚙️ Input Configuration", expanded=True):
        st.markdown("#### 📝 Client Details & Reference Upload")

        # Client name input
        client_name = st.text_input("Enter Client Name (required)", "", key="client_name_coreass")

        # Single unified uploader (matches integrationV3)
        uploaded_file = st.file_uploader(
            "Upload RFP Document",
            type=["pdf", "docx", "xlsx", "pptx"],
            key="rfp_uploader_coreass",
            help="Upload PDF, Word, Excel or PowerPoint reference document.",
        )

        st.markdown("---")

        st.markdown("#### 📋 Select Sections to Generate (in order)")
        st.caption("✨ Tick sections in the order you want them generated. They will be numbered #1, #2, etc.")

        SECTION_LIST = [
            "Executive Summary",
            "About Crave InfoTech",
            "Our Understanding",
            "Project Scope",
            "Solution",
            "Project Approach",
            "Team Structure",
            "Commercials & Payment Terms",
            "Sign Off",
            "Key Assumptions",
        ]

        # Initialize checkbox states in session state (once)
        if "checkbox_states_coreasses" not in st.session_state:
            st.session_state["checkbox_states_coreasses"] = {section: False for section in SECTION_LIST}

        col1, col2 = st.columns(2)
        for i, section in enumerate(SECTION_LIST):
            col = col1 if i < (len(SECTION_LIST) // 2 + len(SECTION_LIST) % 2) else col2
            with col:
                current_state = st.session_state["checkbox_states_coreasses"][section]
                new_state = st.checkbox(section, value=current_state, key=f"chk_coreass_{i}")
                st.session_state["checkbox_states_coreasses"][section] = new_state

        # Build selected sections list in the order they appear in SECTION_LIST
        selected_sections = [s for s in SECTION_LIST if st.session_state["checkbox_states_coreasses"].get(s)]

        # Display selected sections with order numbers
        if selected_sections:
            st.markdown("---")
            st.markdown("### ✅ Selected Sections (in generation order)")
            cols_display = st.columns(min(3, len(selected_sections)))
            for idx, section in enumerate(selected_sections):
                with cols_display[idx % len(cols_display)]:
                    st.markdown(f"**#{idx + 1}** — {section}")

        st.markdown("---")

        # Generate button inside expander (mirrors integrationV3)
        if st.button("⚡ Generate Content"):
            st.session_state.pop("edited_sections", None)

            # Reset editors
            for key in list(st.session_state.keys()):
                if key.startswith("editor_"):
                    st.session_state.pop(key)

            # Allow generation even without RFP
            reference_text = st.session_state.get("reference_text", "")
            if not reference_text:
                reference_text = ""
                st.info("ℹ No RFP uploaded — generating a generic SOW draft.")

            if not selected_sections:
                st.warning("⚠ Please select at least one section to generate.")
            else:
                with st.spinner(f"⏳ Generating {len(selected_sections)} selected sections..."):
                    generated = generate_selected_sections(client, model_name, client_name, selected_sections)

                MASTER_ORDER = SECTION_LIST

                ordered_list = []
                for section_name in MASTER_ORDER:
                    if section_name in generated:
                        content = generated[section_name].strip()
                        ordered_list.append({"title": section_name, "content": content})

                st.session_state["edited_sections"] = ordered_list
                st.success(f"✅ Generated {len(selected_sections)} sections!")


    # Process uploaded reference (if any)
    if uploaded_file and "reference_text" not in st.session_state:
        try:
            name = uploaded_file.name.lower()
            if name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file).fillna("")
                st.session_state["excel_markdown"] = df.to_markdown(index=False)
                st.success("Excel appendix processed and stored for insertion.")
            else:
                raw = extract_text_from_file(uploaded_file)
                st.session_state["reference_text"] = summarize_large_rfp(raw)
                st.success("Uploaded reference processed and used as contextual reference.")
        except Exception:
            st.warning("Could not process uploaded reference. Continuing without it.")

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
            "Solution": "<SOLUTION>",
            "Project Approach": "<DELIVERY_APPROACH>",
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

