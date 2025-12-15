import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document

# Assuming these imports work as in your original code
import concurrent.futures

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
    insert_table_at_placeholder,
    remove_table_by_tag,
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

def generate_executive_summary(client, model_name, reference_text, client_name):
    """Generate ONLY the Executive Summary for an EAM SOW"""

    prompt = f"""
You are a Senior SAP EAM Consultant from Crave InfoTech.

Generate ONLY the Executive Summary section for an SAP EAM Implementation SOW.

CLIENT: {client_name}

REFERENCE TEXT (from RFP):
{reference_text}

STYLE GUIDELINES:
- Tone must be professional, confident, and business-focused.
- Do NOT use markdown, bullets, or bold formatting.
- Write 2–3 paragraphs, each around 5–6 lines.
- Keep content high-level, strategic, and executive-friendly.

CONTENT REQUIREMENTS:
- Provide a high-level overview of the client’s asset management, maintenance, and operational challenges.
- Mention how implementing SAP EAM (or SAP S/4HANA Asset Management) will streamline maintenance processes, improve asset reliability, reduce downtime, optimize inventory, and enhance overall operational efficiency.
- Highlight the alignment with best practices in preventive, predictive, and corrective maintenance.
- Emphasize overall value: lifecycle management, governance, and improved resource utilization.
- Reference that the proposed engagement aims to help the client modernize, simplify, and standardize their asset management operations.

OUTPUT RULES:
- Output ONLY the Executive Summary content.
- No headings, no lists, no extra commentary.
"""
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
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
- Describe Crave InfoTech's expertise in SAP EAM/Asset Management in 2-3 lines
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


def generate_our_understanding_solution(client, model_name, reference_text, client_name, total_interfaces=None):

    prompt = f"""
You are a Senior SAP EAM Consultant from Crave InfoTech.

Generate the "Our Understanding & Solution" section for {client_name}'s SAP EAM Implementation.
DO NOT use bold (**text**) or markdown. Headings must be plain text.

CRITICAL:
- Keep it high-level, business-focused (not technical and not configuration-level).
- Total length should be: 5–7 paragraphs + a short bullet list.
- Each paragraph must be 4–5 lines.
- Do NOT include technical details (no table structures, no system settings, no master data objects like functional locations or equipment).

MANDATORY STRUCTURE:

3.1 Our Understanding
Write 2–3 paragraphs:
- Understanding of the client's current asset management practices and challenges.
- Key challenges such as high unplanned downtime, reactive maintenance, poor spare parts inventory, or lack of mobile accessibility.
- Strategic need for modernization: improved asset performance, standardized maintenance workflows, better cost control.

3.2 Our Proposed Solution
Write 2–3 paragraphs:
- A high-level SAP EAM implementation approach (e.g., S/4HANA Asset Management, SAP PM).
- Why this solution is the right platform for end-to-end asset lifecycle management (planning, scheduling, execution, analysis).
- Business benefits: improved wrench time, reduced MRO inventory, compliance with safety, and increased asset utilization.
- Keep content business-oriented, not technical.

3.3 Challenges They Are Facing
Provide 5–6 bullets:
- High-level operational, organizational, and data challenges.
- Focus on change management, master data cleansing (assets, BOMs), organizational silos, lack of standard processes, and user adoption.

RULES:
- No markdown.
- No bold text.
- No extra headings outside 3.1, 3.2, 3.3.
- Output ONLY the section content.
"""

    response = client.chat.completions.create(
        model= "Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )

    return response.choices[0].message.content.strip()


def generate_project_scope(client, model_name, reference_text, client_name, total_interfaces=None):
    """
    Generate Project Scope with fixed SOW structure (EAM Version)
    """

    prompt = f"""
You are writing the "Project Scope" section for {client_name}'s SAP EAM implementation.

CRITICAL:
- This is a Statement of Work (SOW). Keep the content HIGH-LEVEL.
- DO NOT include technical configuration details (no technical objects, no specific transaction codes).
- NO bold (**text**), NO markdown, NO numbering styles.

Write ONLY these four sub-sections in order:

4.1 Proposed Solution
Write a single high-level paragraph (6–8 lines) describing:
- The proposed SAP EAM implementation approach and key capabilities.
- Coverage of core maintenance processes (Corrective, Preventive, Calibration, Work Clearance Management, Mobile Integration).
- High-level integration with SAP MM for spare parts and inventory.
- The strategic value of standardizing asset lifecycle management processes.

4.2 Deliverables
- List all project deliverables.
- Functional and technical documentation.
- Configuration and design documents.
- Testing evidence (UT/SIT/UAT).
- Training materials and knowledge transfer.
- Production deployment and cutover readiness documents.

4.3 Acceptance Criteria
- Clear criteria for successful SAP EAM implementation.
- Functional completeness and alignment with business requirements.
- Successful testing outcomes (e.g., successful work order creation, completion, settlement).
- User enablement and sign-off.
- Production deployment readiness.

4.4 Out of Scope
- Explicit list of exclusions for this EAM implementation.
- Any modules (e.g., SAP QM, Project Systems), countries, or business units not included.
- Custom development (Z-reports, enhancements) beyond agreed scope.
- Integration components not covered under the EAM program.

RULES:
- 4.1 must be a single high-level paragraph (NOT bullets).
- 4.2, 4.3, 4.4 must use bullet points.
- Keep bullets short and business-friendly.
- DO NOT invent new sections.
- Output ONLY the final content for 4.1 to 4.4.

BEGIN.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()

def generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces=None):

    prompt = f"""
You are a Senior SAP EAM Consultant from Crave InfoTech.

Generate ONLY a short introductory narrative (2–3 lines) for the "Project Delivery Approach" section for {client_name}'s SAP EAM implementation.

CONTEXT:
The detailed project phase table (Preparation, Blueprint, Realization, Testing, Documentation) will already be included in the SOW template. Your output should ONLY provide a high-level introduction explaining the overall delivery methodology.

DO NOT include:
- Headings
- Technical configurations
- Any table content (the template already contains it)

WRITE A 2–3 LINE INTRO PARAGRAPH THAT:
- Explains Crave InfoTech's structured SAP EAM implementation methodology (e.g., using SAP Activate).
- Mentions the focus on asset data governance, collaboration with client maintenance teams, and ensuring process optimization.
- Summarizes how Functional, Technical, and PMO resources jointly execute the project phases.
- Stays high-level, business-focused, and professional.

Output ONLY the short introduction narrative. No bullets, no lists, no extra commentary.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()


def generate_timelines(client, model_name, reference_text, client_name, total_interfaces=None):

    prompt = f"""
Write the full 'Project Timelines & Resources' section for {client_name}'s SAP EAM implementation.
Produce 3 paragraphs of 4–5 lines each. No bullets, no headings, no dates, no numbers.

Guidelines:
- Paragraph 1: Describe the overall project timeline following standard phases (Discover, Prepare, Explore, Realize, Deploy, Run) and how activities progress from design to deployment.
- Paragraph 2: Explain how Crave and client teams coordinate across functional (e.g., maintenance planners), technical, testing, and PMO roles to ensure predictable milestones and governance.
- Paragraph 3: Summarize post-go-live stabilization, hypercare support, knowledge transfer, and how this approach ensures smooth adoption and high asset uptime.

Do NOT mention technical configurations, system interfaces, or middleware.
Use a high-level, business-focused tone.

REFERENCE TEXT:
{reference_text}

Output only the 3 paragraphs.
"""

    response = client.chat.completions.create(
        model=model_name,
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

    prompt = f"""
You are writing the "Key Assumptions" section for {client_name}'s SAP EAM Implementation SOW.

GENERATE:
- 3 to 5 subsections with simple headings (e.g., Client Responsibilities, Master Data Readiness, Infrastructure & Access, Project Governance).
- Each subsection should contain 2–4 short bullet points.
- Bullets must be high-level and project-oriented, not legal or overly detailed.
- Content must be based on standard EAM implementation assumptions.

RULES:
- Add numbering styles like 7.1 for sections.
- Use plain text headings followed by bullets.
- Use real bullets (•).
- Keep it concise and business-focused.
- Do NOT mention technical configurations or system details.
- Output ONLY the assumptions section.

REFERENCE TEXT FROM RFP:
{reference_text}
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
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
        "Executive Summary": lambda: generate_executive_summary(client, model_name, reference_text, client_name),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, reference_text, client_name),
        "Our Understanding & Solution": lambda: generate_our_understanding_solution(client, model_name, reference_text, client_name, total_interfaces),
        "Project Scope": lambda: generate_project_scope(client, model_name, reference_text, client_name, total_interfaces),
        "Project Delivery Approach": lambda: generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces),
        "Project Timelines": lambda: generate_timelines(client, model_name, reference_text, client_name, total_interfaces),
        "Sign Off": lambda: generate_sign_off(client, model_name, client_name),
        "Key Assumptions": lambda: generate_key_assumptions(client, model_name, reference_text, client_name),
    }


    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(section_generators[name]): name
            for name in selected_sections
        }
        
        for future in concurrent.futures.as_completed(futures):
            name = futures[future]
            generated_sections[name] = future.result()

        
    
    return generated_sections


# ========================================
# MAIN APP
# ========================================

def main():
    st.title("🏭 EAM — SOW Generator")
    st.caption("✨ Restructured sections with selective generation")
    
    # Initialize session state (UPDATED KEYS)
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)
    st.session_state.setdefault("client_name_eam", "") # Unique key
    st.session_state.setdefault("uploaded_file_eam", None) # Unique key


    # Azure LLM client
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "gpt-4o-mini"

    st.session_state["llm_client"] = client
    st.session_state["llm_model"] = model_name

    # ========================================
    # Input Configuration + Ordered Section Selection (REPLICATED FROM INTEGRATION)
    # ========================================
    with st.expander("⚙️ Input Configuration", expanded=True):
        
        # --- CLIENT DETAILS & RFP UPLOAD ---
        st.markdown("#### 📝 Client Details & RFP Upload")
        
        # Client name input (UPDATED to use session state key)
        client_name = st.text_input("Enter Client Name (required)", st.session_state["client_name_eam"])
        st.session_state["client_name_eam"] = client_name # Update session state

        # File upload (UPDATED to use session state key and track changes)
        uploaded_file = st.file_uploader(
            "Upload RFP Document",
            type=["pdf", "docx", "xlsx", "pptx"],
            key="rfp_uploader_eam_v2", # Unique key
            help="Upload PDF, Word, Excel or PowerPoint reference document.",
        )
        # Handle file upload change (NEW LOGIC)
        if uploaded_file != st.session_state["uploaded_file_eam"]:
            st.session_state["uploaded_file_eam"] = uploaded_file
            if "reference_text" in st.session_state:
                st.session_state.pop("reference_text") # Clear old reference text if a new file is uploaded
            if uploaded_file: # Only rerun if a file was actually uploaded, not just cleared
                st.experimental_rerun() # Rerun to process file immediately

        st.markdown("---")
        
        # --- SECTION SELECTION ---
        st.markdown("#### 📋 Select Sections to Generate (in order)")
        st.caption("✨ Tick sections in the order you want them generated. They will be numbered #1, #2, etc.")

        # Master list of sections
        SECTION_LIST = [
            "Executive Summary",
            "About Crave InfoTech",
            "Our Understanding & Solution",
            "Project Scope",
            "Project Delivery Approach",
            "Project Timelines",
            "Sign Off",
            "Key Assumptions",
        ]
        
        # Initialize checkbox states in session state (once)
        if "checkbox_states_eam" not in st.session_state: # Unique key
            st.session_state["checkbox_states_eam"] = {section: False for section in SECTION_LIST}
        
        # Display checkboxes in 2 columns with live order tracking
        col1, col2 = st.columns(2)
        
        # Render checkboxes and track state changes (4 items per column)
        for i, section in enumerate(SECTION_LIST):
            col = col1 if i < 4 else col2
            with col:
                # Get current state
                current_state = st.session_state["checkbox_states_eam"][section]
                # Use a unique key for checkboxes in EAM
                new_state = st.checkbox(section, value=current_state, key=f"chk_eam_{section}_v2")
                
                # Update state if changed
                st.session_state["checkbox_states_eam"][section] = new_state
        
        # Build selected sections list in the order they appear in SECTION_LIST
        selected_sections = [s for s in SECTION_LIST if st.session_state["checkbox_states_eam"].get(s)]

        # Display selected sections with order numbers
        if selected_sections:
            st.markdown("---")
            st.markdown("### ✅ Selected Sections (in generation order)")
            cols_display = st.columns(min(3, len(selected_sections)))
            for idx, section in enumerate(selected_sections):
                with cols_display[idx % len(cols_display)]:
                    st.markdown(f"**#{idx + 1}** — {section}")
        
        st.markdown("---")
        
        # ========================================
        # GENERATE BUTTON (INSIDE EXPANDER)
        # ========================================
        
        if st.button("⚡ Generate Content", key="gen_eam"): # Unique key
            st.session_state.pop("edited_sections", None)
            
            # Reset old editor text areas
            for key in list(st.session_state.keys()):
                if key.startswith("editor_"):
                    st.session_state.pop(key)

            # --- PROCESS RFP FILE IF UPLOADED (UPDATED LOGIC) ---
            reference_text = st.session_state.get("reference_text", "")
            
            # Use uploaded_file_eam from session state which is the source of truth after the rerun logic
            uploaded_file_sot = st.session_state["uploaded_file_eam"]
            
            if uploaded_file_sot and not reference_text:
                # Need to process file if not already done
                with st.spinner("Processing RFP..."):
                    raw_text = extract_text_from_file(uploaded_file_sot)
                    if len(raw_text.split()) > 3500:
                        st.session_state["reference_text"] = summarize_large_rfp(client, model_name=model_name, text=raw_text)
                    else:
                        st.session_state["reference_text"] = raw_text
                reference_text = st.session_state["reference_text"]
                st.success(f"✅ Extracted {len(raw_text.split())} words from RFP.")
            elif not uploaded_file_sot:
                reference_text = ""        # Use empty reference
                st.info("ℹ No RFP uploaded — generating a generic SOW draft.")
            
            
            if not selected_sections:
                st.warning("⚠ Please select at least one section to generate.")
            elif not client_name:
                st.warning("⚠ Please enter the Client Name.")
            else:
                # Generate selected sections
                with st.spinner(f"⏳ Generating {len(selected_sections)} selected sections..."):
                    generated_sections = generate_selected_sections(
                        client, model_name, reference_text, client_name, selected_sections
                    )

                MASTER_ORDER = SECTION_LIST

                ordered_list = []

                for section_name in MASTER_ORDER:
                    if section_name in generated_sections:
                        content = re.sub(r"</?[^>]+>", "", generated_sections[section_name]).strip()
                        ordered_list.append({
                            "title": section_name,
                            "content": content
                        })

                st.session_state["edited_sections"] = ordered_list
                st.success(f"✅ Generated {len(selected_sections)} sections!")

    
    # --- RFP Processing (Moved logic to the Generate button block for better control) ---
    # The block below is now only for showing extracted assets if reference_text is already loaded.

    reference_text = st.session_state.get("reference_text", "")

    # ========================================
    # PREVIEW TABS (OUTSIDE INPUT CONFIGURATION)
    # ========================================
    
    if "edited_sections" in st.session_state:
        section_preview_tabs()

    # ========================================
    # DOWNLOAD BUTTON (V2 TEMPLATE)
    # ========================================
    
    if "edited_sections" in st.session_state:
        buffer = io.BytesIO()
        
        template_path = "Template/EAM_Template.docx"
        
        if not os.path.exists(template_path):
            st.warning("⚠ Template not found. Please create EAM_Template.docx")
        
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

        # Insert PPT assets (placeholders remain for completeness, even if EAM doesn't use all)
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
            file_name=f"EAM_SOW_V2_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )