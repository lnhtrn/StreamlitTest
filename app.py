import streamlit as st

pg = st.navigation([
    st.Page("home.py", title="Homepage"), 
    st.Page("mod12.py", title="Module 1&2"), 
    st.Page("mod12_no.py", title="Module 1&2 No Autism"), 
    st.Page("mod3.py", title="Module 3"), 
    st.Page("mod3_no_autism.py", title="Module 3 No Autism"), 
    st.Page("mod4.py", title="Module 4"), 
    # st.Page("gsheet_test.py", title="Testing GSheet"),
])
pg.run()