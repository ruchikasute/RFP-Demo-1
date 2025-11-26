import streamlit as st
from openai import AzureOpenAI
import os, io, re
from datetime import datetime
from docx import Document
from Modules.startup import init
init("Integration")   # <-- single line to initialize everything

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
def extract_section(text, heading, all_headings):
    # remove current heading from next-head list
    filtered = [h for h in all_headings if h != heading]
    escaped = [re.escape(h) for h in filtered]
    next_heads = "|".join(escaped)

    # match heading at start of a line + grab until next heading
    pattern = rf"^{re.escape(heading)}\s*\n(.*?)(?=^({next_heads})\s*$|\Z)"
    match = re.search(pattern, text, re.DOTALL | re.MULTILINE)
    return match.group(1).strip() if match else ""



def generate_all_sections(client, model_name, reference_text, client_name):

    prompt = f"""
You are a Senior SAP Integration Consultant from Crave InfoTech.

Generate a complete SAP PI/PO → SAP Integration Suite Migration Statement of Work (SOW)
following the EXACT structure below.

Client: {client_name}

REFERENCE SOW STYLE GUIDE:
{st.session_state.get("knowledge_text", "")}

Use this as reference:
{reference_text}

REAL INTERFACE COUNT (must be used everywhere):
Total Interfaces Identified: {st.session_state.get("total_interfaces", "UNKNOWN")}

STRICT RULES:
- ALWAYS use the above Total Interfaces value.
- NEVER invent or assume a different number.
- NEVER use example numbers from the knowledge repository.
- DO NOT repeat content already present in the template (Team Structure, Architecture Diagram, Pricing tables, etc.)
- Follow EXACT tag structure.
- No extra text outside tags.

===========================================================
FINAL OUTPUT STRUCTURE (FOLLOW EXACTLY)
===========================================================

Executive Summary
Write a strong executive summary with 300 words covering:
- The purpose of the SOW  
- Overall Migration Approach  
- Summary of deliverables  
- Total interfaces ({st.session_state.get("total_interfaces", "UNKNOWN")})  

About Crave InfoTech  
A concise section describing Crave InfoTech’s capabilities, accelerators,
SAP expertise, and integration migration experience.

Our Understanding of the Client Solution 
3.1 Situation  
- Current PI/PO landscape summary  
- Client’s current middleware challenges  
- Need for modernization / cloud-first alignment  
   
3.2 Objectives  
- What the client wants to achieve through this migration  
- Modernization, reduced TCO, better monitoring, scalability  

3.3 Challenges They Are Facing  
- Technical, operational, compliance, and performance pain points  
- Any migration blockers typically seen in PI/PO → IS projects  

Project Scope
4.1 Proposed Solution  
- Describe our proposed Integration Suite migration solution  
- Key features and capabilities  
- Use of OData, APIs, CPI flows, BTP services, monitoring, etc.  

4.2 Acceptance Criteria
- Clear, verifiable acceptance standards  

4.3 Deliverables 
- All deliverables of this migration project  
- Artifacts, configuration, documentation, testing outputs  

4.4 Out of Scope 
- Explicit list of exclusions  

Solution Details
- Architecture diagram is already in the template (DO NOT regenerate)  
5.1 Bill of Material (BOM)  
- List all required components, licenses, environments (but no pricing)  

Project Approach
(Template includes a 4–5 step image — DO NOT regenerate the image)  
Write a narrative explanation covering:  
- Discovery  
- Assessment  
- Migration  
- Validation  
- Go-Live & Support  
Include how extracted technical tables (integration inventory, complexity, 
classifications, etc.) integrate into these steps.

Team Structure
(Already in template — DO NOT recreate)  
Write a short narrative describing roles & responsibilities that align with
the existing template structure.

Timeline
- The project plan image will be placed by the template (DO NOT generate)  
- Write a narrative summary explaining the phases and duration in general terms.

Commercials & Pricing
(Pricing tables and payment terms are already in the template — DO NOT recreate)  
Write a short explanation on how pricing is structured (without values).

Sign-0FF
[content]

Key Assumptions
Provide a professional list of assumptions applicable to PI/PO → IS migration.

"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role":"user","content":prompt}],
        temperature=0.4
    )

    full_output = response.choices[0].message.content.strip()
    headings = [
    "Executive Summary",
    "About Crave InfoTech",
    "Our Understanding of the Client Solution",
    "Project Scope",
    "Solution Details",
    "Project Approach",
    "Team Structure",
    "Timeline",
    "Commercials & Pricing",
    "Key Assumptions"
]



    # EXTRACT each block
    extracted = {
        "Executive Summary": extract_section(full_output, "Executive Summary", headings),
        "About Crave InfoTech": extract_section(full_output, "About Crave InfoTech", headings),
        "Our Understanding": extract_section(full_output, "Our Understanding of the Client Solution", headings),
        "Project Scope": extract_section(full_output, "Project Scope", headings),
        "Solution Details": extract_section(full_output, "Solution Details", headings),
        "Project Delivery Approach": extract_section(full_output, "Project Approach", headings),
        "Team Structure": extract_section(full_output, "Team Structure", headings),
        "Resource Allocation & Timelines": extract_section(full_output, "Timeline", headings),
        "Commercials": extract_section(full_output, "Commercials & Pricing", headings),
        "Key Assumptions": extract_section(full_output, "Key Assumptions", headings),
    }



    return extracted



def main():
    st.title("🌐 Integration — SOW Generator")

    
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
        # PPT EXTRACTIONS (image, tables, bullets, numeric counts)
        # ---------------------------------------------------------
        if uploaded_file.name.lower().endswith(".pptx"):

            # Slide 17 image
            img = extract_image_from_slide(uploaded_file, 17)
            if img:
                image_blob, ext = img
                st.session_state["slide17_image"] = image_blob
                extracted_items.append("📸 Image (Slide 17)")

            # Table from slide 9
            table_md = extract_table_from_slide(uploaded_file, 9)
            if table_md:
                st.session_state["slide9_table"] = table_md
                extracted_items.append("📊 Table (Slide 9)")

            # Table from slide 8
            table_md_8 = extract_table_from_slide(uploaded_file, 8)
            if table_md_8:
                st.session_state["slide8_table"] = table_md_8
                extracted_items.append("📊 Table (Slide 8)")

            # Slide 7 summary bullets
            slide7_summary = extract_slide7_summary(uploaded_file)
            if slide7_summary:
                st.session_state["slide7_text"] = slide7_summary
                extracted_items.append("📝 Summary bullets (Slide 7)")

            # Slide 5 — total interface count
            total = extract_total_interfaces_from_slide(uploaded_file, slide_number=5)
            if total:
                st.session_state["total_interfaces"] = total
                extracted_items.append(f"🔢 Interface count: {total}")

            # Slide 18 — resources table
            table_md_18 = extract_table_from_slide(uploaded_file, 18)
            if table_md_18:
                st.session_state["slide18_resources_table"] = table_md_18
                extracted_items.append("👥 Resources table (Slide 18)")


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
        template_path = "Template/Integration_Template.docx"
        
        # ---------------------------------------------------------
        # STEP 3 — Build preview sections
        # ---------------------------------------------------------

        titles_and_keys = [
            ("Executive Summary", "Executive Summary"),
            ("About Crave InfoTech", "About Crave InfoTech"),
            ("Our Understanding", "Our Understanding"),
            ("Project Scope", "Project Scope"),
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

            st.session_state["edited_sections"].append(
                {"title": title, "content": content}
            )

 
        pass


        # clear regeneration flag
        st.session_state.pop("regen_success", None)



    # -------------------------------------------------------------
    # SHOW ALL EDITABLE TABS
    # -------------------------------------------------------------
    if "edited_sections" in st.session_state:
        section_preview_tabs()


    # -------------------------------------------------------------
    # DOWNLOAD FINAL SOW DOCX
    # -------------------------------------------------------------
    if "edited_sections" in st.session_state:

        # Generate file only when user clicks Download
        buffer = io.BytesIO()

        template_path = "Template/Integration_Template.docx"
        final_doc = Document(template_path)

        # Basic replacements
        replace_client_name_in_doc(final_doc, client_name)
        replace_submission_date(final_doc)
        doc_no = generate_document_number(client_name)
        insert_document_number(final_doc, "<DOCUMENT_NO>", doc_no)


        placeholder_map = {
            "Executive Summary": "<EXEC_SUMMARY>",
            "About Crave InfoTech": "<ABOUT_CRAVE>",
            "Our Understanding": "<OUR_SOL>",
            "Project Scope": "<PROJECT_SCOPE>",
            "Project Delivery Approach": "<DELIVERY_APPROACH>",
            "Resource Allocation & Timelines": "<RESOURCE_TIMELINE>",
            "Key Assumptions": "<KEY_ASSUMPTIONS>",
        }

        for sec in st.session_state["edited_sections"]:
            title = sec["title"]
            content = sec["content"]
            if title in placeholder_map:
                insert_formatted_text(final_doc, placeholder_map[title], content)



        # for sec in st.session_state["edited_sections"]:
        #     title = sec["title"]
        #     content = sec["content"]
        #     if title in placeholder_map:
        #         insert_formatted_text(final_doc, placeholder_map[title], content)
        #         # insert_plain_preview(final_doc, placeholder_map[title], content)


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

        # 🔥 Save into buffer
        final_doc.save(buffer)
        buffer.seek(0)

        # 🔥 Actual download button
        st.download_button(
            label="📥 Download Final SOW Document",
            data=buffer,
            file_name=f"Integration_SOW_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )



