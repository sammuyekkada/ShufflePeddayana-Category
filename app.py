import streamlit as st
import pandas as pd
from collections import deque
from io import BytesIO

def balanced_round_robin_shuffle(df):
    # Normalize columns
    df.columns = df.columns.str.lower().str.strip()
    grouped = {cat: deque(rows.to_dict('records')) for cat, rows in df.groupby('category')}
    
    category_order = sorted(grouped.keys())  # consistent order
    result = []

    while any(grouped.values()):
        for cat in category_order:
            if grouped[cat]:
                result.append(grouped[cat].popleft())

    return pd.DataFrame(result)

st.title("Even Category Shuffler")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        combined_df = pd.concat(all_sheets.values(), ignore_index=True)

        required_cols = ['name', 'email', 'category']
        combined_df.columns = combined_df.columns.str.lower().str.strip()

        if not all(col in combined_df.columns for col in required_cols):
            st.error(f"Excel must have columns: {required_cols}")
        else:
            shuffled_df = balanced_round_robin_shuffle(combined_df)

            # Create Excel in-memory
            output = BytesIO()
            shuffled_df.to_excel(output, index=False)
            output.seek(0)

            st.success("Rows shuffled in strict category rotation.")
            st.download_button("Download Shuffled Excel", output, file_name="shuffled_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
