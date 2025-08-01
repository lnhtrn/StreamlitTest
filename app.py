import streamlit as st

def user_log_in():
    st.login("google")

def user_log_out():
    st.logout()

# Page navigation
if not st.user.is_logged_in:
    pg = st.navigation([
    ])
else:
    pg = st.navigation([
        st.Page("home.py", title="Homepage"), 
        st.Page("mod12.py", title="Module 1&2", icon="📘"), 
        st.Page("mod12_no.py", title="Module 1&2 No Autism", icon="📘"), 
        st.Page("mod3.py", title="Module 3", icon="📗"), 
        st.Page("mod3_no_autism.py", title="Module 3 No Autism", icon="📗"), 
        st.Page("mod4.py", title="Module 4", icon="📕"), 
        # st.Page("gsheet_test.py", title="Testing Recommendation", icon="🛠️"),
        # st.Page(user_log_out, title="Log out", icon=":material/logout:")
    ])

pg.run()