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

def generate_executive_summary(client, model_name, client_name):
    """Generate ONLY the Executive Summary for a HANA/EE SOW"""

    prompt = f"""
You are a Senior SAP S/4HANA Consultant from Crave InfoTech.

Generate ONLY the Executive Summary section for an SAP S/4HANA Enterprise Extension (EE) or HANA migration SOW.

CLIENT: {client_name}

STYLE GUIDELINES:
- Tone must be professional, confident, and business-focused.
- Do NOT use markdown, bullets, or bold formatting.
- Write 2–3 paragraphs, each around 5–6 lines.
- Keep content high-level, strategic, and executive-friendly.

CONTENT REQUIREMENTS:
- Provide a high-level overview of the client’s need to modernize their ERP landscape (e.g., migrating to S/4HANA or implementing a new Enterprise Extension).
- Mention how this transition/implementation will drive digital transformation, simplify IT, enhance real-time analytics, and improve operational efficiency across core functions (Finance, Logistics, etc.).
- Highlight the strategic value: leveraging the HANA platform for speed, innovation, and future-proofing the organization.
- Emphasize the alignment with industry best practices and a phased approach to manage risk.

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

def generate_about_crave(client, model_name, client_name):
    """Generate ONLY About Crave InfoTech section"""
    prompt = f"""
You are writing the "About Crave InfoTech" section for {client_name}.

REFERENCE STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

INSTRUCTIONS:
- Describe Crave InfoTech's expertise in SAP S/4HANA/HANA migration in 2-3 lines
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


def generate_our_understanding_solution(client, model_name, client_name, total_interfaces=None):

    prompt = f"""
You are a Senior SAP S/4HANA Consultant from Crave InfoTech.

Generate the "Our Understanding & Solution" section for {client_name}'s S/4HANA Implementation/Migration SOW.
DO NOT use bold (**text**) or markdown. Headings must be plain text.

CRITICAL:
- Keep it high-level, strategic, and business-focused.
- Total length should be: 5–7 paragraphs + a short bullet list.
- Each paragraph must be 4–5 lines.
- Do NOT include technical details (no table structures, no system settings, no specific module configurations).

MANDATORY STRUCTURE:

3.1 Our Understanding
Write 2–3 paragraphs:
- Understanding of the client's current challenges (e.g., legacy system limitations, slow reporting, high TCO).
- Need for the S/4HANA transition to simplify data models, enable Fiori, and leverage in-memory computing.
- Strategic need for modernization: process harmonization, global standardization, and digital core adoption.

3.2 Our Proposed Solution
Write 2–3 paragraphs:
- A high-level S/4HANA implementation or migration approach (e.g., Brownfield/Greenfield/Selective Data Transition).
- Why S/4HANA is the right strategic platform for real-time operations and decision-making.
- Business benefits: improved financial closing, optimized supply chain, reduced manual effort, enhanced reporting.
- Keep content business-oriented, not technical.

3.3 Challenges They Are Facing
Provide 5–6 bullets:
- High-level project, organizational, and data challenges.
- Focus on change management, master data cleansing, custom code remediation, business process redesign, and stakeholder alignment.

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


def generate_project_scope(client, model_name, client_name, total_interfaces=None):
    """
    Generate Project Scope with fixed SOW structure (HANA/EE Version)
    """

    prompt = f"""
You are writing the "Project Scope" section for {client_name}'s SAP S/4HANA Implementation/Migration.

CRITICAL:
- This is a Statement of Work (SOW). Keep the content HIGH-LEVEL.
- DO NOT include technical configuration details.
- NO bold (**text**), NO markdown, NO numbering styles.

Write ONLY these four sub-sections in order:

4.1 Proposed Solution
Write a single high-level paragraph (6–8 lines) describing:
- The scope of S/4HANA modules or Enterprise Extensions covered (e.g., Finance, Logistics, specific LOBs).
- The migration or implementation methodology (e.g., Greenfield implementation or Brownfield conversion).
- Focus on process harmonization and adopting standard S/4HANA best practices.
- The strategic value of moving to a unified digital core.

4.2 Deliverables
- List all project deliverables.
- Business Blueprint/Fit-to-Standard documents.
- Configuration and technical design documents.
- Test plans and scripts (UT/SIT/UAT).
- Training materials and knowledge transfer.
- Cutover plan and go-live support documentation.

4.3 Acceptance Criteria
- Clear criteria for successful S/4HANA go-live.
- Functional completeness of scoped processes.
- Successful completion of all test cycles and sign-offs.
- Data migration validation and accuracy.
- User enablement and production readiness sign-off.

4.4 Out of Scope
- Explicit list of exclusions for this S/4HANA project.
- Any non-scoped modules, business units, or countries.
- Custom code development beyond necessary functional gaps.
- Specific non-SAP integrations or third-party systems not explicitly mentioned.

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

def generate_delivery_approach(client, model_name, client_name, total_interfaces=None):

    prompt = f"""
You are a Senior SAP S/4HANA Consultant from Crave InfoTech.

Generate ONLY a short introductory narrative (2–3 lines) for the "Project Delivery Approach" section for {client_name}'s S/4HANA project.

CONTEXT:
The detailed project phase table (e.g., Discover, Prepare, Explore, Realize, Deploy, Run) will already be included in the SOW template. Your output should ONLY provide a high-level introduction explaining the overall delivery methodology (e.g., SAP Activate or a hybrid agile approach).

DO NOT include:
- Headings
- Technical configurations
- Any table content (the template already contains it)

WRITE A 2–3 LINE INTRO PARAGRAPH THAT:
- Explains Crave InfoTech's structured methodology for S/4HANA implementation/migration.
- Mentions the focus on governance, Fit-to-Standard workshops, and collaboration with client Business Process Owners (BPOs).
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


def generate_timelines(client, model_name, client_name, total_interfaces=None):

    prompt = f"""
Write the full 'Project Timelines & Resources' section for {client_name}'s SAP S/4HANA project.
Produce 3 paragraphs of 4–5 lines each. No bullets, no headings, no dates, no numbers.

Guidelines:
- Paragraph 1: Describe the overall project timeline following SAP Activate phases and how activities ensure business readiness for the new digital core.
- Paragraph 2: Explain how Crave and client teams coordinate resources, ensuring involvement of BPOs, IT infrastructure, and change management teams for predictable milestones.
- Paragraph 3: Summarize post-go-live stabilization, hypercare support, knowledge transfer, and how this approach ensures smooth system adoption and realization of business benefits.

Do NOT mention technical configurations, specific transactions, or interface details.
Use a high-level, business-focused tone.

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

def generate_payment_terms(client, model_name, client_name):

    prompt = f"""
Write the 'Payment Terms' section for {client_name}'s SAP S/4HANA Implementation/Migration SOW.

CONTENT:
- Provide a milestone-based payment structure.
- Use 4–6 milestones corresponding to project phases (e.g., Blueprint, Realization completion, Go-Live, Post Go-Live).
- Each milestone must be written as: "• Description — X%".
- Total must equal 100%.

RULES:
- Bullets only.
- No markdown, no tables.
- No headings other than: Payment Terms

Output:
Payment Terms
• ...
• ...
"""
    
    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content.strip()


def generate_key_assumptions(client, model_name, client_name):

    prompt = f"""
You are writing the "Key Assumptions" section for {client_name}'s SAP S/4HANA Implementation/Migration SOW.

GENERATE:
- 3 to 5 subsections with simple headings (e.g., Client Responsibilities, Data Conversion, Infrastructure & Access, Project Governance).
- Each subsection should contain 2–4 short bullet points.
- Bullets must be high-level and project-oriented, not legal or overly detailed.
- Content must be based on standard S/4HANA project assumptions (e.g., client provides BPOs, clean master data available, underlying infrastructure is ready).

RULES:
- Add numbering styles like 7.1 for sections.
- Use plain text headings followed by bullets.
- Use real bullets (•).
- Keep it concise and business-focused.
- Do NOT mention specific technical transactions or tables.
- Output ONLY the assumptions section.
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()


# ========================================
# MAIN GENERATION FUNCTION
# ========================================

def generate_selected_sections(client, model_name, client_name, selected_sections):
    """Generate only the sections that user selected"""
    
    total_interfaces = st.session_state.get("total_interfaces", "UNKNOWN")
    generated_sections = {}
    
    # Map NEW section names to their generation functions
    section_generators = {
        "Executive Summary": lambda: generate_executive_summary(client, model_name, client_name),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, client_name),
        "Our Understanding & Solution": lambda: generate_our_understanding_solution(client, model_name, client_name, total_interfaces),
        "Project Scope": lambda: generate_project_scope(client, model_name, client_name, total_interfaces),
        "Project Delivery Approach": lambda: generate_delivery_approach(client, model_name, client_name, total_interfaces),
        "Project Timelines": lambda: generate_timelines(client, model_name, client_name, total_interfaces),
        "Payment Terms": lambda: generate_payment_terms(client, model_name, client_name),
        "Sign Off": lambda: generate_sign_off(client, model_name, client_name),
        "Key Assumptions": lambda: generate_key_assumptions(client, model_name, client_name),
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
    st.title("🚀 HANA EE — SOW Generator")
    st.caption("✨ Restructured sections with selective generation")
    
    # Initialize session state (UPDATED KEYS)
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)
    st.session_state.setdefault("client_name_hana", "") # Unique key
    st.session_state.setdefault("uploaded_file_hana", None) # Unique key


    # Azure LLM client
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "gpt-4o-mini"

    st.session_state["llm_client"] = client
    st.session_state["llm_model"] = model_name
    
    # Ensure reference_text is initialized or retrieved
    reference_text = st.session_state.get("reference_text", "")


    # ========================================
    # Input Configuration + Ordered Section Selection (REPLICATED)
    # ========================================
    with st.expander("⚙️ Input Configuration", expanded=True):
        
        # --- CLIENT DETAILS & RFP UPLOAD ---
        st.markdown("#### 📝 Client Details & RFP Upload")
        
        # Client name input (UPDATED to use session state key)
        client_name = st.text_input("Enter Client Name (required)", st.session_state["client_name_hana"])
        st.session_state["client_name_hana"] = client_name # Update session state

        # File upload (UPDATED to use session state key and track changes)
        uploaded_file = st.file_uploader(
            "Upload RFP Document",
            type=["pdf", "docx", "xlsx", "pptx"],
            key="rfp_uploader_hana_ee", # Unique key
            help="Upload PDF, Word, Excel or PowerPoint reference document.",
        )
        # Handle file upload change (NEW LOGIC)
        if uploaded_file != st.session_state["uploaded_file_hana"]:
            st.session_state["uploaded_file_hana"] = uploaded_file
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
            "Payment Terms",
            "Sign Off",
            "Key Assumptions",
        ]
        
        # Initialize checkbox states in session state (once)
        if "checkbox_states_hana" not in st.session_state: # Unique key
            st.session_state["checkbox_states_hana"] = {section: False for section in SECTION_LIST}
        
        # Display checkboxes in 2 columns with live order tracking
        col1, col2 = st.columns(2)
        
        # Render checkboxes and track state changes (5 items in col1, 4 items in col2)
        for i, section in enumerate(SECTION_LIST):
            col = col1 if i < 5 else col2
            with col:
                # Get current state
                current_state = st.session_state["checkbox_states_hana"][section]
                # Use a unique key for checkboxes in HANA EE
                new_state = st.checkbox(section, value=current_state, key=f"chk_hana_{section}_v2")
                
                # Update state if changed
                st.session_state["checkbox_states_hana"][section] = new_state
        
        # Build selected sections list in the order they appear in SECTION_LIST
        selected_sections = [s for s in SECTION_LIST if st.session_state["checkbox_states_hana"].get(s)]

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
        
        if st.button("⚡ Generate Content", key="gen_hana"): # Unique key
            st.session_state.pop("edited_sections", None)
            
            # Reset old editor text areas
            for key in list(st.session_state.keys()):
                if key.startswith("editor_"):
                    st.session_state.pop(key)

            # --- PROCESS RFP FILE IF UPLOADED (UPDATED LOGIC) ---
            reference_text = st.session_state.get("reference_text", "")
            
            # Use uploaded_file_hana from session state which is the source of truth after the rerun logic
            uploaded_file_sot = st.session_state["uploaded_file_hana"]
            
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
                        client, model_name, client_name, selected_sections
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

    
    # ========================================
    # PREVIEW TABS
    # ========================================
    
    if "edited_sections" in st.session_state:
        section_preview_tabs()

    # ========================================
    # DOWNLOAD BUTTON (V2 TEMPLATE)
    # ========================================
    
    if "edited_sections" in st.session_state:
        buffer = io.BytesIO()
        
        # NOTE: You'll need to create a NEW template for V2 with updated placeholders
        template_path = "Template/HANA_EE_Template.docx"
        
        # If V2 template doesn't exist yet, fallback to original
        if not os.path.exists(template_path):
            st.warning("⚠ Template not found, please create HANA_EE_Template.docx")
            template_path = "Template/HANA_EE_Template.docx"
        
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
            "Project Timelines": "<TIMELINES>", 
            "Payment Terms": "<PAYMENT_TERMS>",
            "Sign Off": "<SIGN_OFF>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>"
        }

        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)


        final_doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="📥 Download Final SOW Document",
            data=buffer,
            file_name=f"HANA_EE_SOW_V2_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )