import streamlit as st

pg = st.navigation([
    st.Page("mod12.py", title="Module 1&2"), 
    st.Page("mod12_no.py", title="Module 1&2 No Autism"), 
    st.Page("mod3.py", title="Module 3"), 
    st.Page("audio_test.py", title="Transcription Testing"),
])
pg.run()