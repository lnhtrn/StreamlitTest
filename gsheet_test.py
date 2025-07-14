import streamlit as st
import yaml
from streamlit_gsheets import GSheetsConnection

connections = {}

# Scores for sidebar
connections['Scores'] = st.connection(f"mod3_scores", type=GSheetsConnection)
# Read object
df = connections['All'].read(
    ttl="30m",
    usecols=list(range(6)),
    nrows=30,
) 
scores = df.to_dict('records')


# Display data 
yaml_string = yaml.dump(scores, sort_keys=False)
yaml_data = st.code(yaml_string, language=None)