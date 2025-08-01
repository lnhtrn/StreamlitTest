import streamlit as st

# Page navigation
pg = st.navigation([
    st.Page("home.py", title="Homepage"), 
    st.Page("mod12.py", title="Module 1&2", icon="📘"), 
    st.Page("mod12_no.py", title="Module 1&2 No Autism", icon="📘"), 
    st.Page("mod3.py", title="Module 3", icon="📗"), 
    st.Page("mod3_no_autism.py", title="Module 3 No Autism", icon="📗"), 
    st.Page("mod4.py", title="Module 4", icon="📕"), 
    st.Page("gsheet_test.py", title="Testing Recommendation", icon="🛠️"),
])


pg.run()