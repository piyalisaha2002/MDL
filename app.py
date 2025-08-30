import streamlit as st
import pandas as pd
import os 

# file_path = os.path.join(os.path.dirname(__file__), "CI-Extraction.xlsx")
file_path = os.path.join(os.getcwd(), "CI-Extraction.xlsx")

# Load Excel file
# file_path = 'CI-Extraction.xlsx'
df = None
excel_error = None
try:
    df = pd.read_excel(file_path, header=None)
except Exception as e:
    excel_error = str(e)

# Constant marker values
marker_values_input = ["DR", "D", "U", "X,D"]

# Marker colors
marker_colors = {
    "DR": "#808080",  # Grey
    "U": "#FFA500",  # Orange
    "D": "#008000",  # Green
    "X,D": "#008000"  # Green
}

def get_matching_documents(function_name, selected_stages):
    if df is None:
        return [["Excel file could not be loaded: " + excel_error, "", "", ""] + [""] * len(selected_stages)]
    try:
        stage_columns = df.loc[3, 12:18].values.tolist()
        column_headers = df.loc[3, 5:10].tolist()  # Headers for columns F to K
        stage_indices = [12 + stage_columns.index(stage_name) for stage_name in selected_stages if stage_name in stage_columns]
        function_rows = df[df[3].astype(str).str.strip() == function_name].index.tolist()
        if not function_rows:
            return [["Function name not found.", "", "", ""] + [""] * len(selected_stages) + [""] * len(column_headers)]
        
        start_idx = function_rows[0]
        end_idx = start_idx + 1
        while end_idx < len(df):
            next_cell = df.loc[end_idx, 3]
            if pd.isna(next_cell) or str(next_cell).strip().isupper():
                break
            end_idx += 1
        
        function_filtered_df = df.iloc[start_idx + 1:end_idx]
        document_level_df = function_filtered_df[
            function_filtered_df[[0, 1]].apply(lambda row: all(pd.notna(row[i]) and str(row[i]).strip() != '' for i in range(2)), axis=1)
        ]

        rows = []
        for idx, row in document_level_df.iterrows():
            document_name = row[3]
            if pd.notna(document_name) and document_name.strip() != "":
                markers = [row[stage_index] for stage_index in stage_indices]
                if any(marker in marker_values_input for marker in markers):
                    document_row = [document_name, "", ""]
                    # Add data from columns F to K first
                    for col_index in range(5, 11):
                        document_row.append(row[col_index])
                    for stage_index in stage_indices:
                        marker = row[stage_index]
                        if marker in marker_values_input:
                            color = marker_colors.get(marker, "")
                            if color:
                                document_row.append(f'<div style="width: 20px; height: 20px; border-radius: 50%; background-color: {color}; display: block; margin: auto;"></div>')
                            else:
                                document_row.append("")
                        else:
                            document_row.append("")
                    rows.append(document_row)

        if not rows:
            rows = [["No matching documents found.", "", "", ""] + [""] * len(column_headers) + [""] * len(selected_stages)]

        return rows
    except Exception as e:
        return [["An error occurred during document extraction: " + str(e), "", "", ""] + [""] * len(selected_stages) + [""] * len(column_headers)]

# Streamlit UI
st.title("Document Extraction Tool")
st.write("Enter the function name and stage name to get matching documents from the Excel file.")

# Extract function names
if df is not None:
    mask = (
        df[0].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)) &
        df[1].isna() &
        df[2].isna()
    )
    function_names = sorted(set(str(x).strip() for x in df.loc[mask, 3] if pd.notna(x) and str(x).strip() != ""))
else:
    function_names = []

# Add an empty string as the default option
function_names = [""] + function_names

function_name = st.selectbox("Function Name", function_names, index=0)

# Extract stage names
if df is not None:
    stage_names = [str(x) for x in df.loc[3, 12:18].tolist() if pd.notna(x) and str(x).strip() != ""]
else:
    stage_names = []

selected_stages = st.multiselect("Stage Name(s)", stage_names)

# Display results
if st.button("Enter"):
    if function_name and selected_stages:
        rows = get_matching_documents(function_name, selected_stages)
        column_headers = df.loc[3, 5:10].tolist()
        columns = ["SL.No", "Document Number", "Drawing Number"] + column_headers + selected_stages

        # Convert the list of rows to a DataFrame
        result_df = pd.DataFrame(rows, columns=columns)
        
        # Generate HTML for the table to embed CSS for bold headers and auto-fix table
        table_html = result_df.to_html(index=False, escape=False)
        
        # Embedding custom CSS to bold headers and adjust table display
        st.markdown("""
        <style>
        .table {
            width: 100%;
            border-collapse: collapse;
            border: 1px solid black;
        }
        .table thead th {
            font-weight: bold;
        }
        .table td, .table th {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(table_html, unsafe_allow_html=True)
    
    else:
        st.warning("Please select a function name and at least one stage name.")
        





