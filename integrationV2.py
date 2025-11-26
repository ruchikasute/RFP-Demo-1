import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document

# Assuming these imports work as in your original code
from Modules.startup import init
init("Integration")

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
from Modules.llm import regenerate_section_llm
from Modules.preview import section_preview_tabs, md_to_html


# ========================================
# V2 RESTRUCTURED SECTION-SPECIFIC PROMPTS
# ========================================

def generate_executive_summary(client, model_name, reference_text, client_name, total_interfaces):
    """Generate ONLY Executive Summary"""

    prompt = f"""
You are a Senior SAP Integration Consultant from Crave InfoTech.

Generate ONLY the Executive Summary section for the SAP PI/PO → SAP Integration Suite Migration SOW.

Client: {client_name}

REFERENCE STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

RFP REFERENCE:
{reference_text}

CRITICAL DATA:
Total Interfaces: {total_interfaces}

INSTRUCTIONS (STRICT):
- Write **3 short paragraphs** (5–6 lines each)
- High-level business summary only (no complexity %, no approach steps)
- Mention the total interface count ONCE
- Include **1 short bullet list IF it strengthens clarity** (max 3–4 bullets)
- Bullets should summarize key business benefits, not technical details
- Do NOT include detailed migration approach or team structure
- Tone must be professional, confident, and executive-level

Output ONLY the executive summary content. No tags, no extra text.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_about_crave(client, model_name, reference_text, client_name):
    """Generate ONLY About Crave InfoTech section"""
    prompt = f"""
You are writing the "About Crave InfoTech" section for {client_name}.

REFERENCE STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

INSTRUCTIONS:
- Describe Crave InfoTech's expertise in SAP Integration in 2-3 lines
- Highlight migration factory experience
- Professional tone showcasing credibility

Output ONLY the "About Crave InfoTech" content. No tags, no extra text.
"""
    
    response = client.chat.completions.create(
        model="Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_our_understanding_solution(client, model_name, reference_text, client_name, total_interfaces):

    prompt = f"""
You are a Senior SAP Integration Consultant from Crave InfoTech.

Generate the "Our Understanding & Solution" section for {client_name}.
DO NOT use bold (**text**) or markdown. Headings must be plain text.

CRITICAL:
- Keep it high-level, business-focused (not technical and not assessment-style).
- Total length should be: **5–7 paragraphs + a short bullet list**.
- Each paragraph must be **4–5 lines** (not too short).
- No technical jargon like adapter types, Java mappings, credential migration, etc.
- Mention {total_interfaces} interfaces only once.

MANDATORY STRUCTURE:

3.1 Our Understanding
Write **2–3 paragraphs**:
- Understanding of client's integration landscape and transformation goals
- Why modernization is required (cloud readiness, scalability, future architecture)
- Mention the total interfaces ({total_interfaces}) once in a natural way
- Keep it strategic, not technical

3.2 Our Proposed Solution
Write **2–3 paragraphs**:
- High-level migration solution and approach
- Why SAP Integration Suite on BTP is the right strategic platform
- Business benefits: agility, governance, standardization, future readiness
- No low-level technical details

3.3 Challenges They Are Facing
Provide **5–6 bullets**:
- High-level operational and organizational challenges
- Avoid technical blockers
- Focus on change management, testing coordination, data governance, stakeholder alignment, rollout readiness

RULES:
- No markdown
- No bold text
- No extra headings outside 3.1, 3.2, 3.3
- Output ONLY the section content
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )

    return response.choices[0].message.content.strip()


def generate_project_scope(client, model_name, reference_text, client_name, total_interfaces):
    """
    Generate Project Scope with fixed SOW structure.
    """

    prompt = f"""
You are writing the "Project Scope" section for {client_name}'s SAP Integration Suite migration.

CRITICAL:
- This is a Statement of Work (SOW). Keep the content HIGH-LEVEL.
- DO NOT include technical details like file adapters, mappings, Java code, credential migration, etc.
- DO NOT mention complexity categories (Migrate/Adapt/Evaluate).
- DO NOT include numbers except the total interface count {total_interfaces}.
- DO NOT add any paragraphs outside 4.1, 4.2, 4.3, 4.4.
- NO bold (**text**), NO markdown, NO numbering styles.

Write ONLY these four sub-sections in order:

4.1 Proposed Solution
Write a single high-level paragraph (6–8 lines) describing:
- Our high-level migration solution (no technical internals)
- Key features and capabilities of SAP Integration Suite
- Use of OData, APIs, CPI flows, BTP services, monitoring

4.2 Deliverables
- All project deliverables
- Documentation, configuration artifacts, testing evidence, KT session

4.3 Acceptance Criteria
- Clear, verifiable criteria for successful migration

4.4 Out of Scope
- Explicit list of exclusions

RULES:
- 4.1 must be a single high-level paragraph (NOT bullets)
- 4.2, 4.3, 4.4 must use bullet points
- Keep bullets short and business-friendly.
- DO NOT invent new sections.
- DO NOT write assessment-style technical details.
- Output ONLY the final content for 4.1 to 4.4.

BEGIN.
"""

    response = client.chat.completions.create(
        model="Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()


import concurrent.futures
def generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces):

    prompt = f"""
Write ONLY the Project Approach for {client_name}'s SAP PI/PO → SAP Integration Suite migration.

Use this structure:
- 2–3 line intro paragraph
- Then 5 bullets (•):
  • Discovery – 3–4 lines. Mention extracting inventory, complexity and classification tables ({total_interfaces} interfaces).
  • Assessment – 3–4 lines. Mention migrate/adapt/evaluate categories.
  • Migration – 3–4 lines. Describe execution approach.
  • Validation – 3–4 lines. UT/SIT testing support.
  • Go-Live & Support – 3–4 lines. Hypercare, transition.

Rules:
- No headings (no “6.1”, no bold, no markdown)
- No extra sections
- No images
- Output ONLY the intro + bullet points.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_timelines(client, model_name, reference_text, client_name, total_interfaces):
    """
    RENAMED SECTION 6: Project Timelines
    (formerly "Resource Allocation & Timelines")
    """
    prompt = f"""
You are writing the "Project Timelines & Resources" section for {client_name}.

RFP REFERENCE:
{reference_text}

CRITICAL DATA:
Total Interfaces: {total_interfaces}

INSTRUCTIONS:
- Write a SHORT narrative paragraph (3-5 lines ONLY)
- Mention estimated timeline based on interface count
- Reference resource allocation plan
- Keep it concise and high-level
- Do NOT create detailed tables (those are in template)

Output ONLY a brief paragraph. No tags, no extra text.
"""
    
    response = client.chat.completions.create(
        model="Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_sign_off(client, model_name, client_name):
    """Generate ONLY Sign Off section"""
    prompt = f"""
You are writing the "Sign Off" section for {client_name}.

INSTRUCTIONS:
- Write formal agreement statement
- Mention both parties (Crave InfoTech and {client_name})
- Include standard legal language for SOW acceptance
- Keep it brief (2-3 sentences)
- Professional and formal tone

Output ONLY the "Sign Off" content. No tags, no extra text.
"""
    
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_key_assumptions(client, model_name, reference_text, client_name):
    """
    Generate Key Assumptions with EXACT nested numbering:
    1, 1.1, 1.1.1, bullet list, 1.2, bullets, 1.3, bullets, 1.4 paragraph, 1.5 paragraph.
    """

    prompt = f"""
You are writing the "Key Assumptions" section for {client_name}'s
SAP PI/PO → SAP Integration Suite Migration SOW.

REFERENCE:
{reference_text}

STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

INSTRUCTIONS — FOLLOW EXACTLY:

Write the Key Assumptions in THIS STRUCTURE ONLY:

7.1 Dependency on Client

7.1.1 IT and Infrastructure
• 3–4 bullet points

7.1.2 Network and Connectivity
• 2–3 bullet points

7.1.3 Customer responsibility during the implementation
• 3–4 bullet points

7.2 Other Assumptions
• 5–7 bullet points

7.3 Intellectual Property Rights
• 3–4 bullet points

7.4 Limitation of Liability
Write a single paragraph of 4–5 lines.

7.5 General Provisions
Write a single paragraph of 4–5 lines.

MANDATORY RULES:
- Use REAL bullets (•) under sub-topics.
- Do NOT add any extra sections.
- Do NOT add any explanation outside the structure.
- DO NOT rephrase the numbering.
- DO NOT output headings like "Assumptions:" — only the structure above.

Output ONLY the final structured content. No tags, no extra notes.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4
    )

    return response.choices[0].message.content.strip()


# ========================================
# MAIN GENERATION FUNCTION
# ========================================

def generate_selected_sections(client, model_name, reference_text, client_name, selected_sections):
    """Generate only the sections that user selected"""
    
    total_interfaces = st.session_state.get("total_interfaces", "UNKNOWN")
    generated_sections = {}
    
    # Map NEW section names to their generation functions
    section_generators = {
        "Executive Summary": lambda: generate_executive_summary(client, model_name, reference_text, client_name, total_interfaces),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, reference_text, client_name),
        "Our Understanding & Solution": lambda: generate_our_understanding_solution(client, model_name, reference_text, client_name, total_interfaces),
        "Project Scope": lambda: generate_project_scope(client, model_name, reference_text, client_name, total_interfaces),
        "Project Delivery Approach": lambda: generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces),
        "Project Timelines": lambda: generate_timelines(client, model_name, reference_text, client_name, total_interfaces),
        "Sign Off": lambda: generate_sign_off(client, model_name, client_name),
        "Key Assumptions": lambda: generate_key_assumptions(client, model_name, reference_text, client_name),
    }
    
    # Generate only selected sections
    # for section_name in selected_sections:
    #     if section_name in section_generators:
    #         with st.spinner(f"⏳ Generating {section_name}..."):
    #             generated_sections[section_name] = section_generators[section_name]()


    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(section_generators[name]): name
            for name in selected_sections
        }
        
        for future in concurrent.futures.as_completed(futures):
            name = futures[future]
            generated_sections[name] = future.result()

        
    
    return generated_sections

from Modules.docx_table_handler import extract_table_by_tag
# ========================================
# MAIN APP
# ========================================

def main():
    st.title("🌐 Integration — SOW Generator")
    st.caption("✨ Restructured sections with selective generation")
    
    # Initialize session state
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)

    # Client name input
    client_name = st.text_input("Enter Client Name (required)", "")

    # File upload
    uploaded_file = st.file_uploader(
        "Upload RFP Document",
        type=["pdf", "docx", "xlsx", "pptx"],
        key="rfp_uploader",
        help="Upload PDF, Word, Excel or PowerPoint reference document.",
    )

    # Azure LLM client
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "gpt-4o-mini"

    st.session_state["llm_client"] = client
    st.session_state["llm_model"] = model_name

    # Extract and process uploaded file
    if uploaded_file and "reference_text" not in st.session_state:
        raw_text = extract_text_from_file(uploaded_file)
        extracted_items = []
        st.success(f"✅ Extracted {len(raw_text.split())} words")

        # PPT-specific extractions
        if uploaded_file.name.lower().endswith(".pptx"):
            img = extract_image_from_slide(uploaded_file, 17)
            if img:
                image_blob, ext = img
                st.session_state["slide17_image"] = image_blob
                extracted_items.append("📸 Image (Slide 17)")

            table_md = extract_table_from_slide(uploaded_file, 9)
            if table_md:
                st.session_state["slide9_table"] = table_md
                extracted_items.append("📊 Table (Slide 9)")

            table_md_8 = extract_table_from_slide(uploaded_file, 8)
            if table_md_8:
                st.session_state["slide8_table"] = table_md_8
                extracted_items.append("📊 Table (Slide 8)")

            slide7_summary = extract_slide7_summary(uploaded_file)
            if slide7_summary:
                st.session_state["slide7_text"] = slide7_summary
                extracted_items.append("📝 Summary bullets (Slide 7)")

            total = extract_total_interfaces_from_slide(uploaded_file, slide_number=5)
            if total:
                st.session_state["total_interfaces"] = total
                extracted_items.append(f"🔢 Interface count: {total}")

            table_md_18 = extract_table_from_slide(uploaded_file, 18)
            if table_md_18:
                st.session_state["slide18_resources_table"] = table_md_18
                extracted_items.append("👥 Resources table (Slide 18)")

        if extracted_items:
            st.markdown("### 🔎 Extracted assets")
            for item in extracted_items:
                st.markdown(f"- {item}")

        # Store reference text
        if len(raw_text.split()) > 3500:
            st.session_state["reference_text"] = summarize_large_rfp(client, model_name=model_name, text=raw_text)
        else:
            st.session_state["reference_text"] = raw_text

    reference_text = st.session_state.get("reference_text", "")

    # ========================================
    # V2 RESTRUCTURED SECTION SELECTION
    # ========================================
    
    st.markdown("---")
    st.subheader("📋 Select Sections to Generate (Structure)")
    
    col1, col2 = st.columns(2)
    
    with col1:
        exec_summary = st.checkbox("1. Executive Summary", value=True, key="chk_exec")
        about_crave = st.checkbox("2. About Crave InfoTech", value=True, key="chk_crave")
        our_solution = st.checkbox("3. Our Understanding & Solution", value=True, key="chk_solution")
        project_scope = st.checkbox("4. Project Scope", value=True, key="chk_scope")
    
    with col2:
        delivery_approach = st.checkbox("5. Project Delivery Approach", value=True, key="chk_delivery")
        timelines= st.checkbox("6. Project Timelines", value=True, key="chk_timelines")
        sign_off = st.checkbox("7. Sign Off", value=True, key="chk_signoff")
        key_assumptions = st.checkbox("8. Key Assumptions", value=True, key="chk_assumptions")

    # Build selected sections list
    selected_sections = []
    if exec_summary: selected_sections.append("Executive Summary")
    if about_crave: selected_sections.append("About Crave InfoTech")
    if our_solution: selected_sections.append("Our Understanding & Solution")
    if project_scope: selected_sections.append("Project Scope")
    if delivery_approach: selected_sections.append("Project Delivery Approach")
    if timelines: selected_sections.append("Project Timelines")
    if sign_off: selected_sections.append("Sign Off")
    if key_assumptions: selected_sections.append("Key Assumptions")

    st.markdown("---")

    # ========================================
    # GENERATE BUTTON
    # ========================================
    
    if st.button("⚡ Generate Selected Sections"):
        st.session_state.pop("edited_sections", None)
        
        # Reset old editor text areas
        for key in list(st.session_state.keys()):
            if key.startswith("editor_"):
                st.session_state.pop(key)

        if not reference_text:
            st.warning("⚠ Please upload an RFP first.")
            return
        
        if not selected_sections:
            st.warning("⚠ Please select at least one section to generate.")
            return

        # Generate selected sections
        with st.spinner(f"⏳ Generating {len(selected_sections)} selected sections..."):
            generated_sections = generate_selected_sections(
                client, model_name, reference_text, client_name, selected_sections
            )


        MASTER_ORDER = [
            "Executive Summary",
            "About Crave InfoTech",
            "Our Understanding & Solution",
            "Project Scope",
            "Project Delivery Approach",
            "Project Timelines",
            "Sign Off",
            "Key Assumptions",
        ]

        ordered_list = []

        for section_name in MASTER_ORDER:
            if section_name in generated_sections:
                content = generated_sections[section_name]
                content = re.sub(r"</?[^>]+>", "", content).strip()
                ordered_list.append({
                    "title": section_name,
                    "content": content
                })

        st.session_state["edited_sections"] = ordered_list

        
        st.success(f"✅ Generated {len(selected_sections)} sections!")

    # ========================================
    # PREVIEW TABS
    # ========================================
    
    # if "edited_sections" in st.session_state:
    #     # section_preview_tabs()
        
    if "edited_sections" in st.session_state:
        tabs = section_preview_tabs() 
        sections = st.session_state["edited_sections"]

        # sections = st.session_state["edited_sections"]
        # tab_titles = [s["title"] for s in sections]
        # tabs = st.tabs(tab_titles)

        template_path = "Template/Integration_Template.docx"

        def load_table(key, tag):
            if key not in st.session_state:
                st.session_state[key] = extract_table_by_tag(template_path, tag)

        # Load once
        load_table("df_raci", "{{TABLE_RACI}}")
        load_table("df_assessment", "{{TABLE_ASSESSMENT}}")
        load_table("df_staffing", "{{TABLE_STAFFING}}")
        load_table("df_resources", "{{TABLE_RESOURCES}}")
        load_table("df_commercials", "{{TABLE_COMMERCIALS}}")
        load_table("df_milestones", "{{TABLE_MILESTONES}}")

        # for i, tab in enumerate(tabs):
        #     with tab:
        #         sec = sections[i]
        #         title = sec["title"]
            # Add tables below corresponding sections
        for i, tab in enumerate(tabs):
            with tab:
                title = sections[i]["title"]

                if title == "Project Scope":
                    st.markdown("---")
                    st.subheader("📊 Section Tables")

                    with st.expander("RACI Matrix", expanded=True):
                        st.data_editor(st.session_state["df_raci"], key="edit_raci")

                if title == "Project Delivery Approach":
                    st.markdown("---")
                    st.subheader("📊 Section Tables")
                    with st.expander("Migration Assessment", expanded=False):
                        st.data_editor(st.session_state["df_assessment"], key="edit_assessment")


                if title == "Project Timelines":
                    st.markdown("---")
                    st.subheader("📊 Section Tables")

                    # with st.expander("Staffing Plan", expanded=True):
                    #     st.data_editor(st.session_state["df_staffing"], key="edit_staffing")

                    with st.expander("Resource Allocation", expanded=False):
                        st.data_editor(st.session_state["df_resources"], key="edit_resources")

                    # with st.expander("Commercials", expanded=False):
                    #     st.data_editor(st.session_state["df_commercials"], key="edit_commercials")

                    with st.expander("Payment Milestones", expanded=False):
                        st.data_editor(st.session_state["df_milestones"], key="edit_milestones")
            
    # ========================================
    # DOWNLOAD BUTTON (V2 TEMPLATE)
    # ========================================
    
    if "edited_sections" in st.session_state:
        buffer = io.BytesIO()
        
        # NOTE: You'll need to create a NEW template for V2 with updated placeholders
        template_path = "Template/Integration_Template.docx"
        
        # If V2 template doesn't exist yet, fallback to original
        if not os.path.exists(template_path):
            st.warning("⚠ V2 template not found, using original template. Please create Integration_Template_V2.docx")
            template_path = "Template/Integration_Template.docx"
        
        final_doc = Document(template_path)

        # Basic replacements
        replace_client_name_in_doc(final_doc, client_name)
        replace_submission_date(final_doc)
        doc_no = generate_document_number(client_name)
        insert_document_number(final_doc, "<DOCUMENT_NO>", doc_no)

        # V2 RESTRUCTURED placeholder map
        placeholder_map = {
            "Executive Summary": "<EXEC_SUMMARY>",
            "About Crave InfoTech": "<ABOUT_CRAVE>",
            "Our Understanding & Solution": "<OUR_SOL>",  # NEW placeholder
            "Project Scope": "<PROJECT_SCOPE>",
            "Project Delivery Approach": "<DELIVERY_APPROACH>",
            "Project Timelines": "<TIMELINES>",  # RENAMED placeholder
            "Sign Off": "<SIGN_OFF>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>"
        }

        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)

        # Insert PPT assets
        if "slide17_image" in st.session_state:
            insert_image_at_placeholder(final_doc, "<PPT_IMAGE>", st.session_state["slide17_image"])
        if "slide8_table" in st.session_state:
            insert_formatted_text(final_doc, "<ADAPTER_TABLE>", st.session_state["slide8_table"])
        if "slide9_table" in st.session_state:
            insert_formatted_text(final_doc, "<KEY_TABLE>", st.session_state["slide9_table"])
        if "slide7_text" in st.session_state:
            insert_formatted_text(final_doc, "<SLIDE7_TEXT>", st.session_state["slide7_text"])
        if "slide18_resources_table" in st.session_state:
            insert_formatted_text(final_doc, "<RESOURCES_TABLE>", st.session_state["slide18_resources_table"])
        if "total_interfaces" in st.session_state:
            replace_inline_placeholder(final_doc, "<TOTAL_INTERFACES>", st.session_state["total_interfaces"])

        final_doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Download Final SOW Document",
            data=buffer,
            file_name=f"Integration_SOW_V2_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
