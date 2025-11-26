import streamlit as st
from Modules.knowledge import initialize_vector_db, load_knowledge_text

def init(category="Integration"):
    if "vector_db_ready" not in st.session_state:
        initialize_vector_db()
        st.session_state["vector_db_ready"] = True

    # Load category-specific knowledge (only once)
    if "knowledge_text" not in st.session_state:
        st.session_state["knowledge_text"] = load_knowledge_text(
            query=category,
            category=category
        )
