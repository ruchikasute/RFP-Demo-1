
def generate_all_sections(client, model_name, reference_text, client_name):

    prompt = f"""
You are a Senior SAP EAM Consultant from Crave InfoTech.

Client: {client_name}

Use ONLY the factual information from below reference text:

REFERENCE:
{reference_text}

Write the following SOW sections using clean consulting language.
Do NOT generate any content that is not supported by the reference text.

STRICT RULES:
- First line of each section must NOT start with #, ## or ###.
- Use bullets only when necessary.
- No invented technical details.
- Keep wording compact and business-oriented.
- Bullets must use "- ".
- Headings must use only numbered headings (3.1, 4.1 etc.)

=====================================
OUTPUT FORMAT (FOLLOW EXACT TAGS)
=====================================

<EXEC_SUMMARY>
Write 2–3 business paragraphs.

<ABOUT_CRAVE>
Write 2–3 paragraphs about Crave EAM capabilities.

<ABOUT_CLIENT>
Write:
3.1 About {client_name}
(2 paragraphs)
3.2 Business Verticals
(List 2–3 verticals + 4 bullets each from reference)
3.3 Challenges & Objectives
(2 paragraphs)

<PROJECT_SCOPE>
Write:
4.1 Introduction (2–3 paragraphs)
4.2 Scope of Work (bullets)
4.3 Out of Scope
4.4 Deliverables

<PROPOSED_SOLUTION>
Create a table: requirement | feature | description | solution
Extract data ONLY from the reference text.

<DELIVERY_APPROACH>
1–2 paragraphs

<RESOURCE_TIMELINE>
Short paragraph

<SIGN_OFF>
Short closing paragraph

<KEY_ASSUMPTIONS>
Bulleted assumptions

=====================================
IMPORTANT
=====================================
- Output MUST contain EXACTLY these tags.
- Nothing outside tags.
- No markdown headings except numbered ones.
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
        "Proposed Solution": extract_block("PROPOSED_SOLUTION", full_output),
        "Project Delivery Approach": extract_block("DELIVERY_APPROACH", full_output),
        "Resource Allocation & Timelines": extract_block("RESOURCE_TIMELINE", full_output),
        "Sign Off": extract_block("SIGN_OFF", full_output),
        "Key Assumptions": extract_block("KEY_ASSUMPTIONS", full_output),
    }


    return extracted




import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document

# ------------------------------
# IMPORT ALL EXISTING FUNCTIONS
# ------------------------------

from Modules.extractors import (
    extract_text_from_file,
    summarize_large_rfp,
    extract_image_from_slide,
    extract_table_from_slide,
    extract_slide7_summary,
    extract_total_interfaces_from_slide,
    extract_block
)

from Modules.word_insert import (
    insert_formatted_text,
    insert_plain_preview,
    insert_markdown_table_after,
    insert_image_at_placeholder,
    _create_paragraph_after,
)

from Modules.placeholders import (
    replace_client_name_in_doc,
    replace_submission_date,
    replace_inline_placeholder,
    insert_document_number,
    generate_document_number,
)

from Modules.knowledge import load_knowledge_text
from Modules.llm import (
    regenerate_section_llm
)

from Modules.preview import (
    section_preview_tabs,
    md_to_html
)

if "knowledge_text" not in st.session_state:
    st.session_state["knowledge_text"] = load_knowledge_text()


from concurrent.futures import ThreadPoolExecutor
from pptx import Presentation
def main():
    st.title("🌐 EAM — SOW Generator")

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

    uploaded_file = st.file_uploader(
        " ",
        type=["pdf", "docx", "xlsx", "pptx"],
        key="rfp_uploader",
        help="Upload PDF, Word, Excel or PowerPoint reference document.",
        label_visibility="collapsed"
    )

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


    # -------------------------------------------------------------
    # EXTRACT + SUMMARIZE (only once) — CLEAN UI VERSION
    # -------------------------------------------------------------
    if uploaded_file and "reference_text" not in st.session_state:

        raw_text = extract_text_from_file(uploaded_file)

        # Store extracted items for compact UI display
        extracted_items = []

        st.success(f"Extracted {len(raw_text.split())} words")

        
        # ---------------------------------------------------------
        # CLEAN SUMMARY UI (instead of 10 messages)
        # ---------------------------------------------------------
        if extracted_items:
            st.markdown("### 📎 Extracted assets")
            for item in extracted_items:
                st.markdown(f"- {item}")

        # ---------------------------------------------------------
        # Store reference text
        # ---------------------------------------------------------
        if len(raw_text.split()) > 3500:
            st.session_state["reference_text"] = summarize_large_rfp(
                client,
                model_name=model_name,
                text=raw_text
            )
        else:
            st.session_state["reference_text"] = raw_text

    reference_text = st.session_state.get("reference_text", "")

    # -------------------------------------------------------------
    # GENERATE ALL SECTIONS IN PARALLEL
    # -------------------------------------------------------------
    # if st.button("⚡ Generate SOW"):
    if st.button("⚡ Generate SOW"):
        st.session_state.pop("edited_sections", None)

        # Reset old editor text areas
        for key in list(st.session_state.keys()):
            if key.startswith("editor_"):
                st.session_state.pop(key)


        if not reference_text:
            st.warning("⚠ Please upload an RFP first.")
            return

        # LLM call

        with st.spinner("⏳ Generating all SOW sections..."):
            all_sections = generate_all_sections(client, model_name, reference_text, client_name)


        # ---------------------------------------------------------
        # STEP 1 — Build FINAL processed document in memory
        # ---------------------------------------------------------
        template_path = "Template/EAM_Template.docx"
        
        # ---------------------------------------------------------
        # STEP 3 — Build preview sections
        # ---------------------------------------------------------

        titles_and_keys = [
            ("Executive Summary", "Executive Summary"),
            ("About Crave InfoTech", "About Crave InfoTech"),
            ("About Client", "About Client"),
            ("Project Scope", "Project Scope"),
            ("Proposed Solution", "Proposed Solution"),
            ("Project Delivery Approach", "Project Delivery Approach"),
            ("Resource Allocation & Timelines", "Resource Allocation & Timelines"),
            ("Sign Off", "Sign Off"),
            ("Key Assumptions", "Key Assumptions"),
        ]


        st.session_state["edited_sections"] = []

        for title, key in titles_and_keys:
            content = all_sections.get(key, "").strip()
            if not content:
                content = f"(No content generated for '{title}')"

                # 🔥 1. REMOVE TAGS LIKE <ABOUT>, <PROJECT_SCOPE>
            content = re.sub(r"<\/?[^>]+>", "", content).strip()

                # 🔥 2. CLEAN + STANDARDIZE INTO MARKDOWN
            # content = normalize_to_markdown(content)

            st.session_state["edited_sections"].append(
                {"title": title, "content": content}
            )

        # Initialize editor values so textareas show text
        for i, sec in enumerate(st.session_state["edited_sections"]):
            editor_key = f"editor_{i}"
            if editor_key not in st.session_state:
                st.session_state[editor_key] = sec["content"]
            else:
                # if user edited previously, sync back
                sec["content"] = st.session_state[editor_key]

        # clear regeneration flag
        st.session_state.pop("regen_success", None)



    # -------------------------------------------------------------
    # SHOW ALL EDITABLE TABS
    # -------------------------------------------------------------
    if "edited_sections" in st.session_state:
        # section_preview_tabs(st.session_state["edited_sections"])
        section_preview_tabs()


    # -------------------------------------------------------------
    # DOWNLOAD FINAL SOW DOCX
    # -------------------------------------------------------------
    if "edited_sections" in st.session_state:

        # Generate file only when user clicks Download
        buffer = io.BytesIO()

        template_path = "Template/EAM_Template.docx"
        final_doc = Document(template_path)

        # Basic replacements
        replace_client_name_in_doc(final_doc, client_name)
        replace_submission_date(final_doc)
        doc_no = generate_document_number(client_name)
        insert_document_number(final_doc, "<DOCUMENT_NO>", doc_no)

        # Insert all sections using placeholder map
        placeholder_map = {
            "Executive Summary": "<EXEC_SUMMARY>",
            "About Crave InfoTech": "<ABOUT_CRAVE>",
            "About Client": "<ABOUT_CLIENT>",
            "Project Scope": "<PROJECT_SCOPE>",
            "Proposed Solution": "<PROPOSED_SOLUTION>",
            "Project Delivery Approach": "<DELIVERY_APPROACH>",
            "Resource Allocation & Timelines": "<RESOURCE_TIMELINE>",
            "Sign Off": "<SIGN_OFF>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>"
        }



        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)
                # insert_plain_preview(final_doc, placeholder_map[title], content)


       
        # 🔥 Save into buffer
        final_doc.save(buffer)
        buffer.seek(0)

        # 🔥 Actual download button
        st.download_button(
            label="📥 Download Final SOW Document",
            data=buffer,
            file_name=f"EAM_SOW_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
