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
    
    # Initialize session state
    st.session_state.setdefault("llm_client", None)
    st.session_state.setdefault("llm_model", None)

    # Client name input
    client_name = st.text_input("Enter Client Name (required)", "")

    # # File upload
    # uploaded_file = st.file_uploader(
    #     "Upload RFP Document",
    #     type=["pdf", "docx", "xlsx", "pptx"],
    #     key="rfp_uploader",
    #     help="Upload PDF, Word, Excel or PowerPoint reference document.",
    # )

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
    # if uploaded_file and "reference_text" not in st.session_state:
    #     raw_text = extract_text_from_file(uploaded_file)
    #     extracted_items = []
    #     st.success(f"✅ Extracted {len(raw_text.split())} words")

    #     # Store reference text
    #     if len(raw_text.split()) > 3500:
    #         st.session_state["reference_text"] = summarize_large_rfp(client, model_name=model_name, text=raw_text)
    #     else:
    #         st.session_state["reference_text"] = raw_text

    # reference_text = st.session_state.get("reference_text", "")

    # ========================================
    # V2 RESTRUCTURED SECTION SELECTION
    # ========================================
    
    st.markdown("---")
    st.subheader("📋 Select Sections to Generate (V2 Structure)")
    
    col1, col2 = st.columns(2)
    
    with col1:
        exec_summary = st.checkbox("1. Executive Summary", value=True, key="chk_exec")
        about_crave = st.checkbox("2. About Crave InfoTech", value=True, key="chk_crave")
        our_solution = st.checkbox("3. Our Understanding & Solution", value=True, key="chk_solution")
        project_scope = st.checkbox("4. Project Scope", value=True, key="chk_scope")
    
    with col2:
        delivery_approach = st.checkbox("5. Project Delivery Approach", value=True, key="chk_delivery")
        timelines= st.checkbox("6. Project Timelines", value=True, key="chk_timelines")
        payment_terms = st.checkbox("9. Payment Terms", value=True, key="chk_pay")
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
    if payment_terms: selected_sections.append("Payment Terms")
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

        # if not reference_text:
        #     st.warning("⚠ Please upload an RFP first.")
        #     return
        
        if not selected_sections:
            st.warning("⚠ Please select at least one section to generate.")
            return

        # Generate selected sections
        with st.spinner(f"⏳ Generating {len(selected_sections)} selected sections..."):
            generated_sections = generate_selected_sections(
                client, model_name,  client_name, selected_sections
            )


        MASTER_ORDER = [
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

