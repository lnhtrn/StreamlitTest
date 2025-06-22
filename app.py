import streamlit as st

pg = st.navigation({
    st.Page("mod12.py", title="Module 1-2 Report Builder"), 
    st.Page("audio_test.py", title="Transcription Testing"),
})
pg.run()