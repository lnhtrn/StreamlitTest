import streamlit as st

st.set_page_config(
    page_title="Welcome", 
    page_icon="📝",
    initial_sidebar_state="collapsed",
    layout="centered"
)

if not st.user.is_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1]) # Adjust ratios as needed for desired centering
    with col2:
        st.title("Log in to use Report Builder!")
        if st.button("Log in with Google Account"):
            st.login("google")
    st.stop()

# Sidebar after logging in 
st.sidebar.write(f"Welcome, {st.user.name}!")
if st.sidebar.button("Log out"):
    st.logout()


st.title("Welcome to the Report Builder!")
st.markdown("*For authorized use by Bryan R. Harrison, PhD Psychologist, PC only.*")
st.markdown("---")

st.write("Choose a module below to get started:")

st.page_link("mod12.py", label="Module 1&2", icon="📘")
st.page_link("mod12_no.py", label="Module 1&2 No Autism", icon="📘")
st.page_link("mod3.py", label="Module 3", icon="📗")
st.page_link("mod3_no_autism.py", label="Module 3 No Autism", icon="📗")
st.page_link("mod4.py", label="Module 4", icon="📕")


st.markdown("---")
st.markdown("**Disclaimer:**", unsafe_allow_html=True)
st.markdown("""
This application is intended solely for use in support of work product for Bryan R. Harrison, PhD, Psychologist PC. All patient information must be handled in strict compliance with HIPAA regulations to ensure confidentiality.  

Unauthorized access, use, or distribution of this system and its data is strictly prohibited and may violate intellectual property laws and patient confidentiality protections. Misuse of this application may result in legal action.
""")