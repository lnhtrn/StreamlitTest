import streamlit as st

st.set_page_config(page_title="Welcome", layout="centered")

st.title("Welcome to the Report Builder!")
st.write("Choose a module below to get started:")

st.page_link("mod12.py", label="Module 1&2", icon="📘")
st.page_link("mod12_no.py", label="Module 1&2 No Autism", icon="📗")
st.page_link("mod3.py", label="Module 3", icon="📙")
st.page_link("mod3_no_autism.py", label="Module 3 No Autism", icon="📒")
st.page_link("mod4.py", label="Module 4", icon="📕")
