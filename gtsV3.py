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
    """Generate ONLY the Executive Summary for a GTS SOW"""

    prompt = f"""
You are a Senior SAP GTS Consultant from Crave InfoTech.

Generate ONLY the Executive Summary section for an SAP GTS Implementation SOW.

CLIENT: {client_name}

REFERENCE STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

REFERENCE TEXT (from RFP):
{reference_text}

STYLE GUIDELINES:
- Tone must be professional, confident, and business-focused.
- Do NOT use markdown, bullets, or bold formatting.
- Write 2–3 paragraphs, each around 5–6 lines.
- Keep content high-level, strategic, and executive-friendly.

CONTENT REQUIREMENTS:
- Provide a high-level overview of the client’s global trade, compliance, and customs challenges.
- Mention how implementing SAP GTS will streamline compliance, automate trade processes, reduce manual effort, strengthen audit readiness, and enhance global trade visibility.
- Highlight the alignment with global regulatory requirements (export control, sanctioned party screening, customs processes).
- Emphasize overall value: operational efficiency, risk mitigation, governance, and improved cross-border trade execution.
- Reference that the proposed engagement aims to help the client modernize, simplify, and standardize their trade compliance operations.

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
- Describe Crave InfoTech's expertise in SAP GTS in 2-3 lines
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
You are a Senior SAP GTS Consultant from Crave InfoTech.

Generate the "Our Understanding " section for {client_name}.
DO NOT use bold (**text**) or markdown. Headings must be plain text.

CRITICAL:
- Keep it high-level, business-focused (not technical and not configuration-level).
- Total length should be: 5–7 paragraphs + a short bullet list.
- Each paragraph must be 4–5 lines.
- Do NOT include technical details (no table structures, no system settings, no master data objects).
- Do NOT mention interface counts or PI/PO terminology.

MANDATORY STRUCTURE:

3.1 Our Understanding
Write 2–3 paragraphs:
- Understanding of the client's global trade, customs, and compliance landscape.
- Key challenges such as manual processes, compliance risks, screening delays, decentralized data, or lack of audit readiness.
- Strategic need for modernization: automation, compliance governance, global standardization, improved cycle times.

3.2 Our Proposed Solution
Write 2–3 paragraphs:
- A high-level SAP GTS implementation approach.
- Why SAP GTS is the right strategic platform for compliance, customs, and trade automation.
- Business benefits: risk reduction, operational efficiency, automated screening, improved visibility, regulatory alignment.
- Keep content business-oriented, not technical.

3.3 Challenges They Are Facing
Provide 5–6 bullets:
- High-level operational, regulatory, and organizational challenges.
- Avoid technical blockers.
- Focus on change management, regulatory compliance complexities, data governance, master data readiness, stakeholder alignment, and global rollout coordination.

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
    Generate Project Scope with fixed SOW structure (GTS Version)
    """

    prompt = f"""
You are writing the "Project Scope" section for {client_name}'s SAP GTS implementation.

CRITICAL:
- This is a Statement of Work (SOW). Keep the content HIGH-LEVEL.
- DO NOT include technical configuration details (no master data tables, no system fields, no technical objects).
- DO NOT mention PI/PO, interfaces, middleware, or any Integration Suite terminology.
- DO NOT add any paragraphs outside 4.1, 4.2, 4.3, 4.4.
- NO bold (**text**), NO markdown, NO numbering styles.

Write ONLY these four sub-sections in order:

4.1 Proposed Solution
Write a single high-level paragraph (6–8 lines) describing:
- The proposed SAP GTS implementation approach and key capabilities.
- Coverage of compliance processes such as Sanctioned Party Screening, Export Control, Customs Management, and other relevant modules.
- High-level integration with SAP ERP for trade, logistics, and compliance processes.
- The strategic value of standardizing and automating global trade processes.

4.2 Deliverables
- List all project deliverables.
- Functional and technical documentation.
- Configuration and design documents.
- Testing evidence (UT/SIT/UAT).
- Training materials and knowledge transfer.
- Deployment and cutover readiness documents.

4.3 Acceptance Criteria
- Clear criteria for successful SAP GTS implementation.
- Functional completeness and alignment with business requirements.
- Successful testing outcomes.
- User enablement and sign-off.
- Production deployment readiness.

4.4 Out of Scope
- Explicit list of exclusions for this GTS implementation.
- Any modules, countries, or business units not included.
- Custom development beyond agreed scope.
- Integration components not covered under the GTS program.

RULES:
- 4.1 must be a single high-level paragraph (NOT bullets).
- 4.2, 4.3, 4.4 must use bullet points.
- Keep bullets short and business-friendly.
- DO NOT invent new sections.
- DO NOT write system-level technical details.
- Output ONLY the final content for 4.1 to 4.4.

BEGIN.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()

def generate_solution_section(client, model_name, reference_text, client_name):
    """Generate the GTS Solution section following Integration-style structure."""

    knowledge_text = st.session_state.get("knowledge_text", "")
    placeholder = "[[ARCHITECTURE_IMG]]"

    prompt = f"""
You are a Senior SAP GTS Consultant from Crave InfoTech.

Generate Section 5 — Solution for {client_name}'s SAP GTS Implementation SOW.

REFERENCE KNOWLEDGE (style & tone guidelines):
{knowledge_text}

REFERENCE TEXT (from RFP):
{reference_text}

=====================================================
STRUCTURE TO FOLLOW (STRICT)
=====================================================

5.1 Proposed Architecture
Write a 6–8 line high-level paragraph describing:
- Overall SAP GTS architecture and landscape positioning
- Alignment between SAP ERP and GTS for trade compliance processes
- Screening, export control, customs management, and risk management flows
- Governance, data alignment, and regulatory readiness
- High-level integration touchpoints (non-technical)
- Mention that an architecture diagram is referenced

Then on a NEW LINE output ONLY:
{placeholder}

5.2 Bill of Material (BOM)
Write 4–6 bullets (• symbol) listing key SAP GTS components and supporting systems:
• SAP GTS – Compliance Management  
• SAP GTS – Customs Management  
• SAP GTS – Risk Management  
• SAP ERP (ECC or S/4HANA)  
• SAP NetWeaver / GTS Application Server  
• Supporting SAP Fiori / Reporting Tools  

=====================================================
RULES
=====================================================
- No markdown (no **, no #)
- No bold, no italics
- Use only text + bullets (•)
- The placeholder {placeholder} must appear EXACTLY once
- Keep content business-focused, non-technical
- Output ONLY the final section text
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )

    return response.choices[0].message.content.strip()



def generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces=None):

    prompt = f"""
You are a Senior SAP GTS Consultant from Crave InfoTech.

Generate ONLY a short introductory narrative (2–3 lines) for the "Project Approach" section for {client_name}'s SAP GTS implementation.

CONTEXT:
The detailed project phase table (Preparation, Blueprint, Realization, Testing, Documentation) will already be included in the SOW template. Your output should ONLY provide a high-level introduction explaining the overall delivery methodology.

DO NOT include:
- Headings
- Technical configurations
- Any table content (the template already contains it)

WRITE A 2–3 LINE INTRO PARAGRAPH THAT:
- Explains Crave InfoTech's structured SAP GTS implementation methodology.
- Mentions the focus on governance, collaboration with client teams, and ensuring compliance readiness.
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
Write the full 'Project Timelines & Resources' section for {client_name}'s SAP GTS implementation.
Produce 3 paragraphs of 4–5 lines each. No bullets, no headings, no dates, no numbers.

Guidelines:
- Paragraph 1: Describe the overall project timeline following SAP Activate phases (Discover, Prepare, Explore, Realize, Deploy, Run) and how activities progress from design to deployment.
- Paragraph 2: Explain how Crave and client teams coordinate across functional, technical, testing, and PMO roles to ensure predictable milestones and governance.
- Paragraph 3: Summarize post-go-live stabilization, hypercare support, knowledge transfer, and how this approach ensures smooth adoption and compliance readiness.

Do NOT mention PI/PO, Integration Suite, interfaces, or technical configurations.
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
You are writing the "Key Assumptions" section for {client_name}'s SAP GTS Implementation SOW.

GENERATE:
- 3 to 5 subsections with simple headings (e.g., Client Responsibilities, Data Readiness, Infrastructure & Access, Project Governance, Others).
- Each subsection should contain 2–4 short bullet points.
- Bullets must be high-level and project-oriented, not legal or overly detailed.
- Content must be based on standard GTS implementation assumptions and any relevant cues from the RFP.

RULES:
- Add numbering styles like 7.1 for secctions.
- Use plain text headings followed by bullets.
- Use real bullets (•).
- Keep it concise and business-focused.
- Do NOT mention PI/PO, Integration Suite, interfaces, or technical configurations.
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
        "Our Understanding": lambda: generate_our_understanding_solution(client, model_name, reference_text, client_name, total_interfaces),
        "Project Scope": lambda: generate_project_scope(client, model_name, reference_text, client_name, total_interfaces),
        "Solution": lambda: generate_solution_section(client, model_name, reference_text, client_name),   # NEW
        "Project Approach": lambda: generate_delivery_approach(client, model_name, reference_text, client_name, total_interfaces),
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
    st.title("🌐 GTS — SOW Generator")
    st.caption("✨ Restructured sections with selective generation")
    
    # Initialize session state
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)
    st.session_state.setdefault("client_name_gts", "")
    st.session_state.setdefault("uploaded_file_gts", None)


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
        
        # Client name input
        client_name = st.text_input("Enter Client Name (required)", st.session_state["client_name_gts"])
        st.session_state["client_name_gts"] = client_name # Update session state

        # File upload
        uploaded_file = st.file_uploader(
            "Upload RFP Document",
            type=["pdf", "docx", "xlsx", "pptx"],
            key="rfp_uploader_gts_v2",
            help="Upload PDF, Word, Excel or PowerPoint reference document.",
        )
        # Handle file upload change
        if uploaded_file != st.session_state["uploaded_file_gts"]:
            st.session_state["uploaded_file_gts"] = uploaded_file
            if "reference_text" in st.session_state:
                st.session_state.pop("reference_text") # Clear old reference text if a new file is uploaded
            st.rerun() # Rerun to process file immediately

        st.markdown("---")
        
        # --- SECTION SELECTION ---
        st.markdown("#### 📋 Select Sections to Generate (in order)")
        st.caption("✨ Tick sections in the order you want them generated. They will be numbered #1, #2, etc.")

        # Master list of sections
        SECTION_LIST = [
            "Executive Summary",
            "About Crave InfoTech",
            "Our Understanding",
            "Project Scope",
            "Solution", 
            "Project Approach",
            "Project Timelines",
            "Sign Off",
            "Key Assumptions",
        ]
        
        # Initialize checkbox states in session state (once)
        if "checkbox_states_gts" not in st.session_state:
            st.session_state["checkbox_states_gts"] = {section: False for section in SECTION_LIST}
        
        # Display checkboxes in 2 columns with live order tracking
        col1, col2 = st.columns(2)
        
        # Render checkboxes and track state changes
        for i, section in enumerate(SECTION_LIST):
            col = col1 if i < 4 else col2
            with col:
                # Get current state
                current_state = st.session_state["checkbox_states_gts"][section]
                # Use a unique key for checkboxes in GTS
                new_state = st.checkbox(section, value=current_state, key=f"chk_gts_{section}_v2")
                
                # Update state if changed
                st.session_state["checkbox_states_gts"][section] = new_state
        
        # Build selected sections list in the order they appear in SECTION_LIST
        selected_sections = [s for s in SECTION_LIST if st.session_state["checkbox_states_gts"].get(s)]

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
        
        if st.button("⚡ Generate Content"):
            st.session_state.pop("edited_sections", None)
            
            # Reset old editor text areas
            for key in list(st.session_state.keys()):
                if key.startswith("editor_"):
                    st.session_state.pop(key)

            # --- PROCESS RFP FILE IF UPLOADED ---
            reference_text = st.session_state.get("reference_text", "")
            
            if uploaded_file and not reference_text:
                # Need to process file if not already done
                with st.spinner("Processing RFP..."):
                    raw_text = extract_text_from_file(uploaded_file)
                    if len(raw_text.split()) > 3500:
                        st.session_state["reference_text"] = summarize_large_rfp(client, model_name=model_name, text=raw_text)
                    else:
                        st.session_state["reference_text"] = raw_text
                reference_text = st.session_state["reference_text"]
                st.success(f"✅ Extracted {len(raw_text.split())} words from RFP.")
            elif not uploaded_file:
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

                MASTER_ORDER = SECTION_LIST # Use the section list as the master order

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
        
        template_path = "Template/GTS_Template.docx"
        
        if not os.path.exists(template_path):
            st.warning("⚠ Template not found. Please create GTS_Template.docx")
        
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
            "Our Understanding": "<OUR_SOL>",  # NEW placeholder
            "Project Scope": "<PROJECT_SCOPE>",
            "Solution": "<SOLUTION_SEC>", 
            "Project Approach": "<DELIVERY_APPROACH>",
            "Project Timelines": "<TIMELINES>",  # RENAMED placeholder
            "Sign Off": "<SIGN_OFF>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>"
        }

        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)

        # Insert PPT assets (placeholders remain for completeness, even if GTS doesn't use all)
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
            file_name=f"GTS_SOW_V2_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )