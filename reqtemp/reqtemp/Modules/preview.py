import re
import streamlit as st
from Modules.llm import regenerate_section_llm

def md_to_html(md: str) -> str:
    html = md

    # Escape HTML
    html = html.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # Headings
    html = re.sub(r"^###\s+(.*)$", r"<h4>\1</h4>", html, flags=re.MULTILINE)
    html = re.sub(r"^##\s+(.*)$", r"<h3>\1</h3>", html, flags=re.MULTILINE)
    html = re.sub(r"^#\s+(.*)$", r"<h2>\1</h2>", html, flags=re.MULTILINE)

    # Bold / Italic
    html = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html)
    html = re.sub(r"\*(.+?)\*", r"<em>\1</em>", html)

    # Bullet lists
    def bullets_to_ul(text):
        lines = text.splitlines()
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

    # Convert simple markdown tables
    def convert_table(match):
        tbl = match.group(0).strip("\n")
        rows = [r.strip() for r in tbl.splitlines() if r.strip()]
        if len(rows) < 2:
            return tbl

        headers = [h.strip("* ") for h in rows[0].strip("|").split("|")]

        body = rows[2:]
        th = "".join([f'<th style="background:#008FD3;color:#fff;padding:6px;border:1px solid #ddd;">{h}</th>' 
                      for h in headers])

        trs = ""
        for r in body:
            cols = [c.strip() for c in r.strip("|").split("|")]
            tds = "".join([f'<td style="padding:6px;border:1px solid #ddd;background:#f5f9ff;">{c}</td>' 
                           for c in cols])
            trs += f"<tr>{tds}</tr>"

        return f"""
        <table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:14px;">
            <thead><tr>{th}</tr></thead>
            <tbody>{trs}</tbody>
        </table>
        """

    html = re.sub(r"(?:\|[^\n]+\|(?:\n\|[^\n]+\|)+)", convert_table, html)

    # Newlines to paragraphs
    html = html.replace("\n\n", "</p><p>")
    html = html.replace("\n", "<br>")

    return f"<p style='font-family:Calibri;font-size:15px'>{html}</p>"


def section_preview_tabs():
    if "edited_sections" not in st.session_state:
        return

    st.markdown("## 📑 Review, Edit & Preview Sections")

    sections = st.session_state["edited_sections"]
    tabs = st.tabs([sec["title"] for sec in sections])

    # Read flag ONCE
    regen_flag = st.session_state.get("regen_success", None)

    for idx, tab in enumerate(tabs):
        with tab:
            sec = st.session_state["edited_sections"][idx]
            title = sec["title"]

            # ---- Initialize prompt history for this section ----
            st.session_state.setdefault("prompt_history", {})
            st.session_state["prompt_history"].setdefault(title, [])


            col1, col2 = st.columns([1, 3], vertical_alignment="top")

            # ---------------- LEFT (EDIT) ----------------
            with col2:
                st.markdown(f"### ✏️ Edit — {title}")


                editor_key = f"editor_{idx}"

                # If regeneration just happened, update the editor content
                if st.session_state.get("regen_success") == title:
                    st.session_state[editor_key] = sec["content"]


                # Initialize editor state BEFORE rendering the widget
                if editor_key not in st.session_state:
                    st.session_state[editor_key] = sec["content"]

                updated_text = st.text_area(
                    "",
                    key=editor_key,
                    height=380,
                )

                # Keep edited_sections updated
                st.session_state["edited_sections"][idx]["content"] = updated_text

            with col1:
                # # ---- Show Previous Prompts ----
                # if st.session_state["prompt_history"][title]:
                #     st.markdown("### 📝 Previous Instructions")
                #     for p in st.session_state["prompt_history"][title]:
                #         st.markdown(f"- {p}")
                # ---- Show Previous Prompts as chat bubbles ----
                if st.session_state["prompt_history"][title]:
                    st.markdown("### 📝 Previous Instructions")

                    st.markdown(
                            """
                            <div style="
                                max-height: 220px;
                                overflow-y: auto;
                                padding-right: 6px;
                                border: 1px solid #e0e0e0;
                                border-radius: 8px;
                                background: #fafafa;
                            ">
                            """,
                            unsafe_allow_html=True
                        )

                    for p in st.session_state["prompt_history"][title]:
                            st.markdown(
                                f"""
                                <div style="
                                    background:#f1f3f5;
                                    padding:10px 14px;
                                    margin:6px 0;
                                    border-radius:10px;
                                    font-size:14px;
                                    color:#333;
                                    border:1px solid #e0e0e0;
                                ">
                                    {p}
                                </div>
                                """,
                                unsafe_allow_html=True
                            )

                    st.markdown("</div>", unsafe_allow_html=True)


                # User prompt
                prompt_key = f"prompt_{idx}"

                # Clear saved prompt safely BEFORE drawing the widget
                if st.session_state.get("clear_prompt") == prompt_key:
                    st.session_state[prompt_key] = ""
                    st.session_state["clear_prompt"] = None

                user_prompt = st.text_area(
                    f"Instruction for {title}",
                    key=prompt_key,
                    placeholder="e.g., make it more concise...",
                    height=200
                )



                # Regenerate
                if st.button(f"🔁 Regenerate {title}", key=f"regen_{idx}"):
                    if not user_prompt.strip():
                        st.warning("Enter an instruction.")
                    else:
                        with st.spinner("Rewriting section…"):
                            new_content = regenerate_section_llm(
                                client=st.session_state["llm_client"],
                                model_name=st.session_state["llm_model"],
                                section_title=title,
                                original_text=updated_text,
                                user_prompt=user_prompt
                            )

                        # Store everywhere

                        st.session_state["edited_sections"][idx]["content"] = new_content

                        # 🔥🔥 STORE THIS PROMPT IN HISTORY
                        st.session_state["prompt_history"][title].append(user_prompt.strip())
                        st.session_state["clear_prompt"] = f"prompt_{idx}"

                        # st.session_state[editor_key] = new_content
                        st.session_state["regen_success"] = title

                        st.rerun()

            
    # Clear flag AFTER all rendering
    if regen_flag:
        st.session_state["regen_success"] = None

    return tabs

