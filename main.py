import streamlit as st
import integration, coreasses, gts, ai # 👈 This will call your current async generator

if "initialized" not in st.session_state:
    st.session_state.view = "home"
    st.session_state.initialized = True

# -------------------------------------------------------
# 1. PAGE CONFIGURATION
# -------------------------------------------------------
st.set_page_config(page_title="RFP Proposal AI Generator", layout="wide")

# -------------------------------------------------------
# 2. SESSION INITIALIZATION (FIX)
# -------------------------------------------------------
if "view" not in st.session_state:
    st.session_state.view = "home"  # ✅ only once on first load

# -------------------------------------------------------
# 2. CUSTOM CSS
# -------------------------------------------------------
st.markdown("""
<style>
:root {
    --primary-blue: #1A75E0;
    --light-blue-bg: #EAF3FF;
}

/* Header */
.main-header {
    text-align: center;
    color: #000;
    font-size: 3em;
    font-weight: 800;
    padding-top: 20px;
    padding-bottom: 5px;
}
.highlight-text { color: var(--primary-blue); }
.sub-tagline {
    text-align: center;
    color: #555;
    font-size: 1.1em;
    padding-bottom: 40px;
}

/* Buttons */
.button-box {
    background: #F9F9F9;
    border-radius: 15px;
    padding: 40px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
}
div.stButton > button {
    background-color: var(--primary-blue);
    color: white;
    border-radius: 10px;
    border: none;
    font-size: 1.1em;
    font-weight: 600;
    padding: 15px 20px;
    transition: all 0.2s ease-in-out;
}
div.stButton > button:hover {
    background-color: #145CB0;
    transform: scale(1.05);
}

/* Back Button */
.back-btn {
    display: flex;
    justify-content: center;
    margin-top: 30px;
}
div[data-testid="stButton"][data-key="back_home"] > button {
    background-color: white !important;
    color: var(--primary-blue) !important;
    border: 2px solid var(--primary-blue);
    border-radius: 8px;
    padding: 10px 25px;
    font-weight: 600;
}
div[data-testid="stButton"][data-key="back_home"] > button:hover {
    background-color: var(--light-blue-bg) !important;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------
# 3. NAVIGATION STATE
# -------------------------------------------------------
if "view" not in st.session_state:
    st.session_state.view = "home"

# -------------------------------------------------------
# 4. HOME PAGE
# -------------------------------------------------------
if st.session_state.view == "home":
    st.markdown("<div class='main-header'>Automate Your <span class='highlight-text'>Proposal Response</span></div>", unsafe_allow_html=True)
    st.markdown("<p class='sub-tagline'>Respond to RFPs in minutes with AI-driven content generation.</p>", unsafe_allow_html=True)

    st.markdown("<div class='button-box'>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center; color:#333;'>Select a Module to Continue</h3>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 6, 1])

    # st.markdown("</div>", unsafe_allow_html=True)
    with col2:
        # Add spacing between button columns
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        
        # Make 2 rows of 4 buttons each for balance
        row1 = st.columns(4, gap="large")
        with row1[0]:
            if st.button("🚀 Integration", use_container_width=True):
                st.session_state.view = "integration"
                st.rerun()
        with row1[1]:
            if st.button("💼 CoreAssess.AI", use_container_width=True):
                st.session_state.view = "coreasses"
                st.rerun()
        with row1[2]:
            if st.button("🌍 GTS", use_container_width=True):
                st.session_state.view = "gts"
                st.rerun()
        with row1[3]:
            if st.button("🧠 AI", use_container_width=True):
                st.session_state.view = "ai"
                st.rerun()

        st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)  # extra vertical spacing

        row2 = st.columns(4, gap="large")
        with row2[0]:
            if st.button("🏗️ EAM", use_container_width=True):
                st.session_state.view = "eam"
                st.rerun()
        with row2[1]:
            if st.button("💾 HANA EE", use_container_width=True):
                st.session_state.view = "hana"
                st.rerun()
        with row2[2]:
            if st.button("📊 BW Modernization", use_container_width=True):
                st.session_state.view = "hana"
                st.rerun()
        with row2[3]:
            if st.button("💻 App Development", use_container_width=True):
                st.session_state.view = "app"
                st.rerun()

# -------------------------------------------------------
# 5. INTEGRATION MODULE (your RFP app)
# -------------------------------------------------------
elif st.session_state.view == "integration":
    integration.main()  # 👈 Runs your async RFP generator
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 6. CORE ASSESSMENT MODULE
# -------------------------------------------------------
elif st.session_state.view == "coreasses":
    coreasses.main()
    # st.subheader("💼 Core Assessment Module")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------
# 7. GTS MODULE
# -------------------------------------------------------
elif st.session_state.view == "gts":
    gts.main()
    # st.subheader("💼 Core Assessment Module")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 8. AI
# -------------------------------------------------------
elif st.session_state.view == "ai":
    ai.main()
    # st.subheader("💼 Core Assessment Module")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 9. EAM
# -------------------------------------------------------
elif st.session_state.view == "eam":
    st.subheader("Will comming soon")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 10. HANA EE
# -------------------------------------------------------
elif st.session_state.view == "hana":
    st.subheader("Will comming soon")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 11. BW Modernization
# -------------------------------------------------------
elif st.session_state.view == "hana":
    st.subheader("Will comming soon")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------
# 11. App Development
# -------------------------------------------------------
elif st.session_state.view == "app":
    st.subheader("Will comming soon")
    st.markdown("<div class='back-btn'>", unsafe_allow_html=True)
    if st.button("⬅ Back to Home", key="back_home"):
        st.session_state.view = "home"
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)