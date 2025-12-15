

import re
import streamlit as st
from Modules.llm import regenerate_section_llm

# (md_to_html function remains unchanged)
def md_to_html(md: str) -> str:
    # ... (function body as provided in the prompt)
    import re
    html = md or ""

    # -----------------------------------
    # 1) Normalize line endings
    # -----------------------------------
    html = html.replace("\r\n", "\n").replace("\r", "\n")

    # -----------------------------------
    # 2) Extract placeholders temporarily
    # -----------------------------------
    placeholders = {}
    def _ph_repl(m):
        idx = len(placeholders)
        token = f"__PH_{idx}__"
        placeholders[token] = m.group(0)     # like [[PROJECT_PLAN_IMG]]
        return token

    html = re.sub(r"\[\[[^\]]+\]\]", _ph_repl, html)

    # -----------------------------------
    # 3) Escape HTML (safe: placeholders are protected)
    # -----------------------------------
    html = (html.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;"))

    # -----------------------------------
    # 4) Headings (markdown + numbered SOW)
    # -----------------------------------
    html = re.sub(r"^###\s+(.*)$", r"<h4>\1</h4>", html, flags=re.MULTILINE)
    html = re.sub(r"^##\s+(.*)$", r"<h3>\1</h3>", html, flags=re.MULTILINE)
    html = re.sub(r"^#\s+(.*)$",  r"<h2>\1</h2>", html, flags=re.MULTILINE)
    html = re.sub(r"^####\s+(.*)$", r"<h5>\1</h5>", html, flags=re.MULTILINE)

    # SOW style headings (e.g., 4.1 Title)
    html = re.sub(r"^(\d+\.\d+)\s+(.*)$",
                  r"<h3>\1&nbsp;&nbsp;\2</h3>",
                  html,
                  flags=re.MULTILINE)

    # Apply MS Word-like inline CSS
    html = html.replace(
        "<h2>",
        "<h2 style='margin-top:12px;margin-bottom:10px;font-size:26px;"
        "font-weight:600;color:#004b8d;font-family:Calibri;'>"
    )
    html = html.replace(
        "<h3>",
        "<h3 style='margin-top:2px;margin-bottom:2px;font-size:22px;"
        "font-weight:600;color:#005fa3;font-family:Calibri;'>"
    )
    html = html.replace(
        "<h4>",
        "<h4 style='margin-top:8px;margin-bottom:4px;font-size:18px;"
        "font-weight:600;color:#006bb3;font-family:Calibri;'>"
    )
    html = html.replace(
    "<h5>",
    "<h5 style='margin-top:6px;margin-bottom:4px;font-size:16px;"
    "font-weight:600;color:#0077cc;font-family:Calibri;'>"
    )


    # -----------------------------------
    # 5) Bold / Italic
    # -----------------------------------
    html = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html)
    html = re.sub(r"\*(.+?)\*",   r"<em>\1</em>", html)

    # -----------------------------------
    # 6) Markdown bullet lists → <ul><li>
    # -----------------------------------
    def bullets_to_ul(text):
        lines = text.split("\n")
        out, in_list = [], False
        for ln in lines:
            m = re.match(r"^\s*[-•*]\s+(.*)$", ln)
            if m:
                if not in_list:
                    out.append("<ul>")
                    in_list = True
                out.append(f"<li>{m.group(1)}</li>")
            else:
                if in_list:
                    out.append("</ul>")
                    in_list = False
                out.append(ln)
        if in_list:
            out.append("</ul>")
        return "\n".join(out)

    html = bullets_to_ul(html)
    html = re.sub(r"</li>\s*<br\s*/?>\s*", "</li>", html)

    # -----------------------------------
    # 7) Markdown tables → HTML tables
    # -----------------------------------
    def convert_table(match):
        tbl = match.group(0).strip("\n")
        rows = [r.strip() for r in tbl.splitlines() if r.strip()]
        if len(rows) < 2:
            return tbl

        headers = [h.strip("* ") for h in rows[0].strip("|").split("|")]
        body = rows[2:]

        th = "".join([
            f'<th style="background:#008FD3;color:#fff;padding:6px;'
            f'border:1px solid #ddd;">{h}</th>'
            for h in headers
        ])

        trs = ""
        for r in body:
            cols = [c.strip() for c in r.strip("|").split("|")]
            tds = "".join([
                f'<td style="padding:6px;border:1px solid #ddd;'
                f'background:#f5f9ff;">{c}</td>'
                for c in cols
            ])
            trs += f"<tr>{tds}</tr>"

        return f"""
        <table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:14px;">
            <thead><tr>{th}</tr></thead>
            <tbody>{trs}</tbody>
        </table>
        """

    html = re.sub(r"(?:\|[^\n]+\|(?:\n\|[^\n]+\|)+)",
                  convert_table,
                  html)

    # -----------------------------------
    # NEW FIX — ensure 1 blank line between paragraphs
    # -----------------------------------
    # NEW FIX — ensure 1 blank line between paragraphs
    html = re.sub(r"\n\s*\n", "\n\n", html)

    # 8) Newline normalization
    # html = re.sub(r"\n\s*\n\s*(?:\n\s*)+", "\n\n", html)
    # html = re.sub(r"(?<!__PH_\d__)\n\n(?!__PH_\d__)", "<br>", html)
    # html = re.sub(r"(?<!__PH_\d__)\n(?!__PH_\d__)", "<br>", html)

    # -----------------------------------
    # NEW FIX — proper paragraph spacing
    # -----------------------------------

    # Step 1 — Collapse 2+ blank lines to exactly TWO
    html = re.sub(r"\n\s*\n+", "\n\n", html)

    # Step 2 — Convert paragraph breaks to <br><br>
    html = html.replace("\n\n", "<br><br>")

    # Step 3 — Convert single newlines to <br>
    html = html.replace("\n", "<br>")

    # # ⭐ FIX — Reduce spacing ONLY when next item is NORMAL TEXT, not placeholders
    # html = re.sub(
    #     r"(?<=</h[2-4]>)<br>\s*(?!<br>|\[\[)",
    #     "<br>",
    #     html
    # )
    # -----------------------------------
    # 9) FIX spacing around headings (CORRECT PLACE)
    # -----------------------------------
    # html = re.sub(r"(?:<br>\s*){2,}(?=<h[2-4])", "<br>", html)
    # html = re.sub(r"(?<=</h[2-4]>)(?:\s*<br>){2,}", "<br>", html)

    # -----------------------------------
    # 10) Bullet-list spacing cleanup
    # -----------------------------------
    html = re.sub(r"(?:<br>\s*){2,}(?=<ul>)", "<br>", html)
    html = re.sub(r"(?<=</ul>)(?:\s*<br>){2,}", "<br>", html)
    html = re.sub(r"(?<=</ul>)(?:\s*<br>){2,}(?=<h3)", "<br>", html)

    # -----------------------------------
    # 11) Restore placeholders with spacing
    # -----------------------------------

    # def restore_ph(token):
    #     ph = placeholders[token]
    #     return f"<br><br>{ph}<br><br>"
    def restore_ph(token):
        return placeholders[token]


    for token in placeholders:
        html = html.replace(token, restore_ph(token))
    
    # Add consistent spacing around placeholders
    html = re.sub(r"(\[\[[^\]]+\]\])", r"<br>\1<br>", html)



    # for token in placeholders:
    #     html = html.replace(token, restore_ph(token))

    # -----------------------------------
    # 12) FINAL FIX — collapse excessive spacing between placeholder and next heading
    # -----------------------------------
    html = re.sub(r"(?:<br>\s*){2,}(?=<h[2-4])", "<br>", html)


    return html


def section_preview_tabs():
    if "edited_sections" not in st.session_state:
        return

    st.markdown("## 📑 Preview Sections")

    # --- START OF NEW CSS INJECTION FOR HOLLOW ORANGE BOXED TABS ---
    st.markdown(
        """
        <style>
        /* This targets the entire tab container block to clean up any unwanted borders */
        div[data-testid="stTabs"] {
            border-bottom: none !important;
            padding-bottom: 0 !important;
            margin-bottom: 0 !important;
        }
        
        /* This targets the individual tab buttons (the headers) */
        button[data-testid^="stTab"] {
            /* HOLLOW STYLE: transparent background with orange border */
            background-color: transparent;
            border: 2px solid #ff7f00; /* Orange border */
            border-radius: 5px; /* Slightly rounded corners */
            padding: 5px 12px; /* Adjusted padding */
            margin-right: 10px; 
            margin-bottom: 5px; 
            color: #ff7f00; /* Orange text color */
            font-weight: 500;
        }
        
        /* Style the active tab button (SOLID ORANGE) */
        button[data-testid^="stTab"][aria-selected="true"] {
            background-color: #ff7f00; /* Solid orange background for active tab */
            color: white; /* White text for active tab */
            border-color: #ff7f00;
            margin-bottom: 0; /* Remove margin so it looks connected to the content */
            font-weight: 700;
        }
        
        /* Ensure the content area itself sits flush with the active tab */
        div[data-testid="stTabs"] > div[data-testid="stHorizontalBlock"] {
             padding-top: 15px;
        }
        
        /* Attempt to fix the double tab issue by targeting the first tab block container only (may vary by Streamlit version) */
        /* This is a common Streamlit tab header selector */
        div.stTabs [data-baseweb="tab-list"] {
            flex-wrap: nowrap; /* Prevent wrapping if too many tabs */
        }
        
        /* Hide the duplicate set if it is in a subsequent block (highly specific guess) */
        /* Comment this out if it causes problems, but it's a common trick to fix duplications */
        div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"]:nth-child(2) > div[data-testid="stTabs"] {
            display: none;
        }

        </style>
        """,
        unsafe_allow_html=True,
    )
    # --- END OF NEW CSS INJECTION ---

    sections = st.session_state["edited_sections"]

    # Read flag ONCE
    regen_flag = st.session_state.get("regen_success", None)

    # Create the list of tab titles, adding a success emoji if just regenerated
    tab_titles = []
    for sec in sections:
        title = sec["title"]
        # Use a contrasting color/style for the checkmark to stand out against the orange box
        if regen_flag and title == regen_flag:
            tab_titles.append(f"🟢 {title}")
        else:
            tab_titles.append(title)
            
    # Simple navigation removed: use plain section tabs instead

    tabs = st.tabs(tab_titles)

    for idx, tab in enumerate(tabs):
        with tab:
            sec = st.session_state["edited_sections"][idx]
            title = sec["title"]

            # Removed numeric section header to reduce vertical space

            # ---- Initialize prompt history for this section ----
            st.session_state.setdefault("prompt_history", {})
            st.session_state["prompt_history"].setdefault(title, [])


            col1, col2 = st.columns([1, 3], vertical_alignment="top")

            # LEFT COLUMN: Previous prompts (scrollable) + instruction box
            with col1:
                # Previous Instructions (plain list)
                if st.session_state["prompt_history"][title]:
                    st.markdown("#### Previous Instructions")
                    for p in st.session_state["prompt_history"][title]:
                        st.markdown(f"- {p}")

                # User prompt area sits below the history and has fixed height ~200
                prompt_key = f"prompt_{idx}"
                if st.session_state.get("clear_prompt") == prompt_key:
                    st.session_state[prompt_key] = ""
                    st.session_state["clear_prompt"] = None

                user_prompt = st.text_area(
                    f"Instruction for {title}",
                    key=prompt_key,
                    placeholder="Add your comment here and click submit to edit content",
                    height=200,
                )

                if st.button("🔁 Submit", key=f"submit_btn_{title}"):
                    if not user_prompt.strip():
                        st.warning("Enter an instruction.")
                    else:
                        with st.spinner("Rewriting section…"):
                        #     new_content = regenerate_section_llm(
                        #         client=st.session_state["llm_client"],
                        #         model_name=st.session_state["llm_model"],
                        #         section_title=title,
                        #         original_text=sec["content"],
                        #         user_prompt=user_prompt,
                        #     )

                        # st.session_state["edited_sections"][idx]["content"] = new_content
                        
                        # SAFE ORIGINAL TEXT SELECTION (prevents KeyError)
                            original_text = sec.get("content", sec.get("preview", ""))

                            # Call LLM
                            new_content = regenerate_section_llm(
                                client=st.session_state["llm_client"],
                                model_name=st.session_state["llm_model"],
                                section_title=title,
                                original_text=original_text,
                                user_prompt=user_prompt,
                            )

                        # Store results safely
                        st.session_state["edited_sections"][idx]["content"] = new_content
                        st.session_state["edited_sections"][idx]["preview"] = new_content
                        st.session_state["edited_sections"][idx]["final"] = new_content   # <-- CRITICAL FIX


                        st.session_state["prompt_history"][title].append(user_prompt.strip())
                        st.session_state["clear_prompt"] = f"prompt_{idx}"
                        st.session_state["regen_success"] = title
                        st.rerun()

            # RIGHT COLUMN: Render preview box (keeps existing styling)
            with col2:
                # inject the global bullet CSS ONCE
                st.markdown(
                    """
                    <style>
                    ul {
                        margin-top: 0 !important;
                        margin-bottom: 0 !important;
                        padding-top: 0 !important;
                        padding-bottom: 0 !important;
                    }
                    li {
                        margin-top: 0 !important;
                        margin-bottom: 0 !important;
                        padding-top: 0 !important;
                        padding-bottom: 0 !important;
                        line-height: 1.1 !important;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

                # Plain preview container with a simple border so content isn't floating
                st.markdown(
                    f"""
                    <div style="padding:16px;font-family:Calibri, sans-serif;font-size:16px;line-height:1.55;text-align:justify;height:460px;overflow-y:auto;border:1px solid #e0e0e0;border-radius:6px;background:#ffffff;">
                        {md_to_html(sec["preview"]) }
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            
    # Clear flag AFTER all rendering
    if regen_flag:
        st.session_state["regen_success"] = None

    return tabs