import streamlit as st
import yaml
from streamlit_gsheets import GSheetsConnection


# Set up side bar
def clear_my_cache():
    st.cache_data.clear()

st.markdown("**After editing dropdown options, please reload data using the button below to update within the form.**")
st.button('Reload Dropdown Data', on_click=clear_my_cache)


connections = {}

# Scores for sidebar
connections['Scores'] = st.connection(f"mod3_scores", type=GSheetsConnection)
# Read object
df = connections['Scores'].read(
    ttl="30m",
    usecols=list(range(6)),
    nrows=30,
) 
scores = df.to_dict('records')

# Display data 
yaml_string = yaml.dump(scores, sort_keys=False)
yaml_data = st.code(yaml_string, language=None)