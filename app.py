import streamlit as st

pg = st.navigation([
    st.Page("webpages/mod12.py", title="Module 1&2"), 
    st.Page("webpages/mod12_no.py", title="Module 1&2 No Autism"), 
    st.Page("webpages/audio_test.py", title="Transcription Testing"),
])
pg.run()