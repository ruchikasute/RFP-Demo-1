import streamlit as st
import coreassesV3, aiV2 ,eamV3, appdevV2, hana_eeV2, integrationV4, gtsV3, bwV2# 👈 This will call your current async generator

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
    --primary-blue: #FF7A00;
    --light-blue-bg: #FFF2E6;
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
    background-color: #CC6300;
    transform: scale(1.05);
}

/* BACK TO HOME button */
div[data-testid="stButton"][data-key="back_home"] > button {
    background-color: white !important;
    color: var(--primary-blue) !important;
    border: 2px solid var(--primary-blue);
    border-radius: 8px;
    padding: 10px 25px;
    font-weight: 600;
}
div[data-testid="stButton"][data-key="back_home"] > button:hover {
    background-color: #FFE1C4 !important;
}

/* --- TEXT-ONLY BACK BUTTON (Integration) --- */

/* Make ONLY the 'back_integration' button plain text */
button[kind="primary"][data-testid*="back_integration"] {
    background: none !important;
    border: none !important;
    color: #666 !important;
    padding: 0 !important;
    margin: 0 0 15px 0 !important;
    font-size: 16px !important;
    font-weight: 400 !important;
    box-shadow: none !important;
    cursor: pointer !important;
    width: auto !important;
    transform: none !important;
}

/* Hover effect */
button[kind="primary"][data-testid*="back_integration"]:hover {
    color: #000 !important;
    text-decoration: underline !important;
    background: none !important;
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
    st.markdown(
        "<div class='main-header'>"
        "Automate Your <span class='highlight-text'>SOW Response</span>"
        "</div>",
        unsafe_allow_html=True
    )


    st.markdown("<p class='sub-tagline'>Generate SOWs quickly and consistently using guided templates.</p>", unsafe_allow_html=True)

    # st.markdown("<div class='button-box'>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center; color:#333;'>Select a Domain to Continue</h3>", unsafe_allow_html=True)

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
                st.session_state.view = "hana_ee"
                st.rerun()
        with row2[2]:
            if st.button("📊 BW Modernization", use_container_width=True):
                st.session_state.view = "bw"
                st.rerun()
        with row2[3]:
            if st.button("💻 App Development", use_container_width=True):
                st.session_state.view = "appdev"
                st.rerun()

# -------------------------------------------------------
# 5. INTEGRATION MODULE (your RFP app)
# -------------------------------------------------------

elif st.session_state.view == "integration":

    # TEXT-ONLY BACK BUTTON
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()

    integrationV4.main()


# -------------------------------------------------------
# 6. CORE ASSESSMENT MODULE
# -------------------------------------------------------
elif st.session_state.view == "coreasses":
        # TEXT-ONLY BACK BUTTON
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    coreassesV3.main()



# -------------------------------------------------------
# 7. GTS MODULE
# -------------------------------------------------------
elif st.session_state.view == "gts":
        # TEXT-ONLY BACK BUTTON
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    gtsV3.main()


# -------------------------------------------------------
# 8. AI
# -------------------------------------------------------
elif st.session_state.view == "ai":
        # TEXT-ONLY BACK BUTTON
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    aiV2.main()
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
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    eamV3.main()

# -------------------------------------------------------
# 10. HANA EE
# -------------------------------------------------------
elif st.session_state.view == "hana_ee":
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    hana_eeV2.main()

# -------------------------------------------------------
# 11. BW Modernization
# -------------------------------------------------------
elif st.session_state.view == "bw":
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    bwV2.main()

# -------------------------------------------------------
# 11. App Development
# -------------------------------------------------------
elif st.session_state.view == "appdev":
    if st.button("⬅ Back", key="back_integration"):
        st.session_state.view = "home"
        st.rerun()
    appdevV2.main()
