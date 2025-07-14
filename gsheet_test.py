import streamlit as st
import yaml
from streamlit_gsheets import GSheetsConnection


# Set up side bar
def clear_my_cache():
    st.cache_data.clear()

def get_abbreviation(test_name):
    # Split on en dash (\u2013)
    main_title = test_name.split('\u2013')[0].strip()
    
    # Split into words and get uppercase initials
    abbreviation = ''.join(word[0] for word in main_title.split() if word[0].isupper())
    
    return abbreviation

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
score_list = df.to_dict('records')

# Process data
optional = {}
check_scores = {}

for test in score_list:
    test_name = test["Test name"]
    abbr = get_abbreviation(test_name)
    optional[abbr] = {}

    optional[abbr]["Test name"] = test_name
    all_lines = []
    all_items = {}
    print(f"\nTest: {test_name}")
    
    for i in range(5):
        line_key = f"Line {i}"
        line_value = test.get(line_key)
        if line_value and line_value != ".nan":
            all_lines.append([])
            items = [item.strip() for item in line_value.split(",")]
            for item in items:
                bold = "(bold)" in item
                item_name = item.replace("(bold)", "").strip()
                # write_item(item_name, bold=bold)
                all_lines[i].append((item_name, bold))
                all_items[item_name] = 0
    
    optional[abbr]["Lines"] = all_lines
    optional[abbr]["All items"] = all_items

# Display data 
yaml_string = yaml.dump(optional, sort_keys=False)
yaml_data = st.code(yaml_string, language=None)