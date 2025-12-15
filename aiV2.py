import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document

# Assuming these imports work as in your original code

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
    """Generate ONLY the Executive Summary for a AI SOW"""

    prompt = f"""
You are a Senior AI Consultant from Crave InfoTech.

Generate ONLY the Executive Summary section for an AI Implementation SOW for {client_name}.

GUIDELINES:
- 2–3 paragraphs, each 5–6 lines.
- No markdown, no bullets, no headings.
- Tone: business-focused, high-level, confident.

CONTENT TO INCLUDE:
- A high-level overview of the client’s operational challenges and need for AI-driven modernization.
- How an AI-led solution can automate processes, improve decision-making, reduce manual effort, and increase accuracy.
- Benefits such as efficiency, governance, transparency, predictive insights, and digital transformation.
- How the proposed engagement supports innovation and long-term scalability.

Output ONLY the Executive Summary.
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
- Describe Crave InfoTech's expertise in SAP AI in 2-3 lines
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


def generate_our_understanding_solution(client, model_name,  client_name, total_interfaces=None):

    prompt = f"""
You are writing 'Our Understanding & Solution' for {client_name}'s AI Implementation SOW.

STRUCTURE:
3.1 Our Understanding — 2–3 paragraphs describing:
- The client's business challenges.
- Need for automation, data intelligence, process simplification.
- Current pain points like manual workflows, low visibility, inconsistent data, or slow decision cycles.

3.2 Our Proposed Solution — 2–3 paragraphs describing:
- The proposed AI approach (automation, ML models, LLM-based assistants, process intelligence).
- Expected business value: improved efficiency, accuracy, compliance, agility, and decision insights.
- Scalable architecture, responsible AI principles, and collaborative delivery.

3.3 Challenges They Are Facing — 5–6 bullets:
- High-level operational and organizational challenges.
- No technical issues.

RULES:
- No markdown.
- Only plain text.
- No technical configuration details.
- Use real bullets.

Output the content for 3.1, 3.2, 3.3.
"""


    response = client.chat.completions.create(
        model= "Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )

    return response.choices[0].message.content.strip()


def generate_project_scope(client, model_name,  client_name, total_interfaces=None):
    """
    Generate Project Scope with fixed SOW structure (AI Version)
    """

    prompt = f"""
Write the 'Project Scope' section for {client_name}'s AI Implementation SOW.

STRUCTURE (MANDATORY):
4.1 Proposed Solution — one paragraph, 6–8 lines.
4.2 Deliverables — bullet list.
4.3 Acceptance Criteria — bullet list.
4.4 Out of Scope — bullet list.

GUIDELINES:
- High-level, business-oriented.
- No markdown, no numbering styles.
- No technical configurations, no datasets, no model parameters.

Output the four subsections ONLY.
IMPORTANT:
- After each subsection heading (4.1, 4.2, 4.3, 4.4), insert a NEW LINE.
- For example:
  4.1 Proposed Solution
  <paragraph>

  4.2 Deliverables
  • bullet1
  • bullet2

Do NOT place the paragraph on the same line as the heading.

"""


    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.3
    )

    return response.choices[0].message.content.strip()

import concurrent.futures
def generate_delivery_approach(client, model_name, client_name, total_interfaces=None):

    prompt = f"""
Generate the 'Project Delivery Approach' section for {client_name}'s AI Implementation SOW.

STRUCTURE:
Phase 1: Planning & Discovery — one paragraph.
Phase 2: Design & Data Assessment — one paragraph.
Phase 3: Model Development & Configuration — one paragraph.
Phase 4: Validation & UAT — one paragraph.
Phase 5: Deployment & Go-Live — one paragraph.

RULES:
- No markdown, no bullets.
- Each phase should be a clean paragraph (4–5 lines).
- High-level, business-focused.
- Generic AI language: automation, ML models, data pipelines, LLM integration, testing, governance.

Output all 5 phases with headings.
"""


    response = client.chat.completions.create(
        model="Codetest",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_timelines(client, model_name,  client_name, total_interfaces=None):

    prompt = f"""
Write the 'Project Timelines & Resources' narrative for {client_name}'s AI Implementation SOW.

STRUCTURE:
- Three paragraphs of 4–5 lines each.
- No bullets, no dates, no numbers.

CONTENT:
- Overview of phases from discovery to deployment.
- Collaboration between business SMEs, technical teams, AI consultants, QA teams, and PMO.
- Post-go-live support, hypercare, knowledge transfer, and adoption readiness.

Output only the three paragraphs.
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
        model="Codetest",
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )
    return response.choices[0].message.content.strip()

def generate_payment_terms(client, model_name, client_name):

    prompt = f"""
Write the 'Payment Terms' section for {client_name}'s AI Implementation SOW.

CONTENT:
- Provide a milestone-based payment structure.
- Use 4–6 milestones.
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
Generate the 'Key Assumptions' section for {client_name}'s AI Implementation SOW.

STRUCTURE:
7.1 Client Responsibilities
• 3–4 bullets

7.2 Data Readiness
• 2–4 bullets

7.3 Infrastructure & Access
• 2–4 bullets

7.4 Governance & Stakeholder Alignment
• 2–4 bullets

7.5 Other Assumptions
• 2–4 bullets

RULES:
- Use plain text headings exactly as shown.
- Use real bullets (•).
- No markdown, no extra text, no paragraphs outside bullets.
- Keep assumptions high-level and business-focused.
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
        "Executive Summary": lambda: generate_executive_summary(client, model_name,  client_name),
        "About Crave InfoTech": lambda: generate_about_crave(client, model_name, client_name),
        "Our Understanding & Solution": lambda: generate_our_understanding_solution(client, model_name, client_name, total_interfaces),
        "Project Scope": lambda: generate_project_scope(client, model_name,  client_name, total_interfaces),
        "Project Delivery Approach": lambda: generate_delivery_approach(client, model_name,  client_name, total_interfaces),
        "Project Timelines": lambda: generate_timelines(client, model_name, client_name, total_interfaces),
        "Payment Terms": lambda: generate_payment_terms(client, model_name, client_name),
        "Sign Off": lambda: generate_sign_off(client, model_name, client_name),
        "Key Assumptions": lambda: generate_key_assumptions(client, model_name,  client_name),
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
    st.title("🌐 AI — SOW Generator")
    st.caption("✨ Restructured sections with selective generation")
    
    # Initialize session state (UPDATED)
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)
    st.session_state.setdefault("client_name_ai", "")
    st.session_state.setdefault("uploaded_file_ai", None)

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
    # Input Configuration + Ordered Section Selection (REPLICATED FROM INTEGRATION)
    # ========================================
    with st.expander("⚙️ Input Configuration", expanded=True):
        
        # --- CLIENT DETAILS & RFP UPLOAD ---
        st.markdown("#### 📝 Client Details & RFP Upload")
        
        # Client name input (UPDATED to use session state key)
        client_name = st.text_input("Enter Client Name (required)", st.session_state["client_name_ai"])
        st.session_state["client_name_ai"] = client_name # Update session state

        # File upload (UPDATED to use session state key and track changes)
        uploaded_file = st.file_uploader(
            "Upload RFP Document",
            type=["pdf", "docx", "xlsx", "pptx"],
            key="rfp_uploader_ai",
            help="Upload PDF, Word, Excel or PowerPoint reference document.",
        )
        # Handle file upload change (NEW LOGIC)
        if uploaded_file != st.session_state["uploaded_file_ai"]:
            st.session_state["uploaded_file_ai"] = uploaded_file
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
        if "checkbox_states_ai" not in st.session_state:
            st.session_state["checkbox_states_ai"] = {section: False for section in SECTION_LIST}
        
        # Display checkboxes in 2 columns with live order tracking
        col1, col2 = st.columns(2)
        
        # Render checkboxes and track state changes (5 items in col1, 4 items in col2)
        for i, section in enumerate(SECTION_LIST):
            col = col1 if i < 5 else col2 
            with col:
                # Get current state
                current_state = st.session_state["checkbox_states_ai"][section]
                # Use a unique key for checkboxes in AI
                new_state = st.checkbox(section, value=current_state, key=f"chk_ai_{section}_v2")
                
                # Update state if changed
                st.session_state["checkbox_states_ai"][section] = new_state
        
        # Build selected sections list in the order they appear in SECTION_LIST
        selected_sections = [s for s in SECTION_LIST if st.session_state["checkbox_states_ai"].get(s)]

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
        
        if st.button("⚡ Generate Content", key="gen_ai"): 
            st.session_state.pop("edited_sections", None)
            
            # Reset old editor text areas
            for key in list(st.session_state.keys()):
                if key.startswith("editor_"):
                    st.session_state.pop(key)

            # --- PROCESS RFP FILE IF UPLOADED (UPDATED LOGIC) ---
            reference_text = st.session_state.get("reference_text", "")
            
            # Use uploaded_file_ai from session state which is the source of truth after the rerun logic
            uploaded_file_sot = st.session_state["uploaded_file_ai"]
            
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
        template_path = "Template/AI_Template.docx"
        
        # If V2 template doesn't exist yet, fallback to original
        if not os.path.exists(template_path):
            st.warning("⚠ V2 template not found, using original template. Please create AI_Template_V2.docx")
            template_path = "Template/AI_Template.docx"
        
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
            file_name=f"AI_SOW_V2_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )