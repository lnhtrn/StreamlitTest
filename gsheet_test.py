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
scores = {}
check_scores = {}

for test in score_list:
    test_name = test["Test name"]
    abbr = get_abbreviation(test_name)
    scores[abbr] = {}

    scores[abbr]["Test name"] = test_name
    all_lines = []
    all_items = {}
    print(f"\nTest: {test_name}")
    
    for i in range(5):
        line_key = f"Line {i}"
        line_value = test.get(line_key)
        if line_value and str(line_value) != "nan":
            all_lines.append([])
            items = [item.strip() for item in line_value.split(",")]
            for item in items:
                bold = "(bold)" in item
                item_name = item.replace("(bold)", "").strip()
                # write_item(item_name, bold=bold)
                all_lines[i].append((item_name, bold))
                all_items[item_name] = 0
    
    scores[abbr]["Lines"] = all_lines
    scores[abbr]["All items"] = all_items

# # Display data 
# yaml_string = yaml.dump(scores, sort_keys=False)
# yaml_data = st.code(yaml_string, language=None)

with st.sidebar:
    st.markdown("**After editing dropdown options, please reload data using the button below to update within the form.**")
    st.link_button("Edit Score Options", st.secrets['mod3_scores'])
    st.button('Reload Dropdown Data', on_click=clear_my_cache)
    
    # Checkbox
    st.markdown("**Check to include score in the form:** Scores to report:")
    scq_result = st.checkbox("Social Communication Questionnaire (SCQ) - Lifetime Form")
    teacher_srs_eval = st.checkbox("Teacher's SRS Scores")
    teacher_vineland_eval = st.checkbox("Teacher's Vineland Adaptive Behavior Scales")
    
    for item in scores:
        check_scores[item] = st.checkbox(scores[item]["Test name"])


data = {}

with st.form("BasicInfo"):  
    #################################################
    st.header("Patient's data")

    data["{{Patient Age}}"] = st.number_input("Patient's Age", 0, 100)
    data['{{Patient age unit}}'] = st.radio(
        "Year/month?",
        ("year", "month")
    )

    #################################################
    # Score section
    for test in check_scores:
        if check_scores[test]:
            st.header(scores[test]["Test name"])
            st.markdown(f"*Skip this section if there is no {test} Score*")

            scores[test]["Test Date"] = st.date_input(f"{test} Test Date").strftime("%m/%Y")
            
            for item in scores[test]["All items"]:
                scores[test]["All items"][item] = st.text_input(item)    

    
    submit = st.form_submit_button('Submit')

if submit:
    # Display data 
    yaml_string = yaml.dump(data, sort_keys=False)
    yaml_string = yaml_string + '\n' + yaml.dump(scores, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)
    
