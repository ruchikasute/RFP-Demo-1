import streamlit as st
import pandas as pd
import time
import re
from openai import AzureOpenAI
import os

# ============================================================
# SECTION 1 — Helper Setup
# ============================================================

def call_llm(prompt, client, model_name):
    """Call the Azure OpenAI LLM."""
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"


def format_steps(steps_text):
    """Convert numbered steps into HTML list."""
    if not steps_text:
        return ""
    lines = [line.strip() for line in steps_text.split("\n") if line.strip()]
    cleaned = [re.sub(r'^\d+[\.\)]\s*', '', l) for l in lines]
    return "<ol>" + "".join([f"<li>{line}</li>" for line in cleaned]) + "</ol>"


def build_prompt(row_dict):
    """Build LLM prompt for per-object analysis."""
    content = "\n".join([f"{k}: {v}" for k, v in row_dict.items()])
    return f"""
You are analyzing an SAP object for modernization.

Task:
1. Write one-sentence **Issue**.
2. Write exactly 5 concise **Key Modernization Steps** (numbered 1–5).

Row details:
{content}

Format:
Issue: <one sentence>
Key Modernization Steps:
1. <step 1>
2. <step 2>
3. <step 3>
4. <step 4>
5. <step 5>
"""


def process_file(df, client, model_name):
    """Process each ABAP object and generate Issue + Steps."""
    results = []
    progress = st.progress(0)

    for i, idx in enumerate(df.index):
        row = df.loc[idx].to_dict()
        text = call_llm(build_prompt(row), client, model_name)

        issue, steps = "", ""
        if "Key Modernization Steps:" in text:
            parts = text.split("Key Modernization Steps:")
            issue = parts[0].replace("Issue:", "").strip()
            steps = parts[1].strip()
        else:
            issue = text.strip()

        # ✅ Include Recommended Approach dynamically
        approach_value = ""
        for k in row.keys():
            if "approach" in k.lower():
                approach_value = row[k]
                break

        results.append({
            "Object Name": row.get("Object Name", ""),
            "Issue": f"<b>{issue}</b>",
            "Recommended Approach": approach_value,
            "Key Modernization Steps": format_steps(steps)
        })

        progress.progress((i + 1) / len(df))
        time.sleep(0.2)

    return pd.DataFrame(results)


def generate_sections(df, client, model_name):
    """Generate executive summary, standard approach & pricing sections."""
    df.columns = [c.strip().lower() for c in df.columns]

    # Find approach column
    approach_col = None
    for col in df.columns:
        if "approach" in col:
            approach_col = col
            break

    total = len(df)
    if approach_col:
        df[approach_col] = df[approach_col].astype(str).str.strip().str.lower()
        onstack = df[approach_col].str.contains("on[- ]?stack", case=False, na=False).sum()
        sidebyside = df[approach_col].str.contains("side[- ]?by[- ]?side", case=False, na=False).sum()
        retire = df[approach_col].str.contains("retire", case=False, na=False).sum()
    else:
        onstack = sidebyside = retire = 0

    issues_text = "; ".join(df.get("issue", df.columns[0]).astype(str).tolist()[:10])

    prompt = f"""
You are an SAP modernization consultant.
Prepare an executive summary based on:

Data Summary:
- Total Objects: {total}
- On-Stack: {onstack}
- Side-by-Side: {sidebyside}
- Retire: {retire}

Sample Issues:
{issues_text}

Write **three interconnected sections** in markdown:

### 1. Analysis
Summarize total objects analyzed, approach distribution, and key findings.

### 2. Standard Approach
Describe modernization strategy (On-Stack, Side-by-Side, Retire) based on patterns.

### 3. Pricing and Timeline
Show a short professional 3-phase table with realistic effort and cost.
"""
    return call_llm(prompt, client, model_name)


# ============================================================
# SECTION 2 — MAIN FUNCTION
# ============================================================
def main():
    st.title("🌐 Core Assess.ai Proposal")

    with st.expander("About this Assessment"):
        st.markdown("""
        **Use Case — Clean Core Assessment (CoreAssess.AI)**  
        Problem: Large volumes of custom ABAP code create technical debt, upgrade risks, and maintenance burden.  
        Solution: CoreAssess.AI ingests ABAP object metadata and produces a prioritized assessment — issue statement, recommended modernization approach, and concise modernization steps per object.  
        Input: Excel list of ABAP objects.  
        Output: Executive summary, prioritized object list, and modernization recommendations.
        """)

    uploaded = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("Please upload an Excel file to continue.")
        return

    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"❌ Could not read the file. Error: {e}")
        return

    st.success(f"✅ File `{uploaded.name}` loaded successfully!")
    st.subheader("Preview of Uploaded Data (first 10 rows)")
    st.dataframe(df.head(10), use_container_width=True)

    # Azure OpenAI setup
    client = AzureOpenAI(
        azure_endpoint=os.getenv("AZURE_OPENAI_FRFP_ENDPOINT"),
        api_key=os.getenv("AZURE_OPENAI_FRFP_KEY"),
        api_version=os.getenv("AZURE_OPENAI_FRFP_VERSION")
    )
    model_name = "codetest"

    if st.button("⚡ Generate Assessment Report"):
        final_df = process_file(df, client, model_name)
        st.subheader("📋 Final Object Assessment Table")

        # --- KPI METRICS ---
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Objects", len(final_df))

        # --- Detect 'Approach' column robustly ---
        def normalize_column(col_name: str) -> str:
            return (
                str(col_name)
                .replace("\xa0", " ")
                .replace("\n", " ")
                .replace("\r", " ")
                .strip()
                .lower()
            )

        approach_col = None
        for col in final_df.columns:
            clean_col = normalize_column(col)
            if "approach" in clean_col:
                approach_col = col
                break

        if approach_col:
            df_temp = final_df.copy()
            df_temp[approach_col] = (
                df_temp[approach_col]
                .astype(str)
                .str.lower()
                .str.strip()
                .str.replace(r"[^a-z\- ]", "", regex=True)
            )

            # --- Count categories ---
            onstack_count = df_temp[approach_col].str.contains("on[- ]?stack", case=False, na=False).sum()
            sidebyside_count = df_temp[approach_col].str.contains("side[- ]?by[- ]?side", case=False, na=False).sum()
            retire_count = df_temp[approach_col].str.contains("retire", case=False, na=False).sum()

            col2.metric("🟦 On-Stack", onstack_count)
            col3.metric("🟧 Side-by-Side", sidebyside_count)
            col4.metric("🔴 Retire", retire_count)

            # --- Chart visualization ---

        else:
            st.warning("⚠️ No column containing 'Approach' found in Excel.")
            col2.metric("On-Stack", 0)
            col3.metric("Side-by-Side", 0)
            col4.metric("Retire", 0)

        # --- TABLE PREVIEW ---
        st.markdown("""
            <style>
            .dataframe td {
                white-space: normal !important;
                word-wrap: break-word !important;
                max-width: 500px;
                vertical-align: top !important;
            }
            </style>
        """, unsafe_allow_html=True)

        st.markdown(final_df.to_html(index=False, escape=False), unsafe_allow_html=True)

        csv = final_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download Final Assessment Table",
            data=csv,
            file_name="Final_Assessment_Table.csv",
            mime="text/csv"
        )

        # --- Generate LLM Summary ---
        st.markdown("---")
        st.subheader("📊 Generating Analysis, Approach, and Pricing...")
        sections_text = generate_sections(final_df, client, model_name)
        st.markdown(sections_text, unsafe_allow_html=True)
