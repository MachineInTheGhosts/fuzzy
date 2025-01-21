import streamlit as st
import pandas as pd
from rapidfuzz import process
import os

st.title("Fuzzy Matching App")

st.info("This is a simple app to match names in excel files using fuzzy matching. "
        "It expects two excel files, one with a Name column and the other with a Name and an ID column."
        "It's just dead simple POC mostly just to demo Streamlit."
        "If something like this could be useful in some guise, "
        "let me know and we can add features.")

uploaded_file_1 = st.file_uploader("Upload Names Spreadsheet", type=["xlsx"])
uploaded_file_2 = st.file_uploader("Upload Names and IDs Spreadsheet", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    names_df = pd.read_excel(uploaded_file_1)
    ids_df = pd.read_excel(uploaded_file_2)

    ids_df_names = ids_df["Name"].tolist()
    matches = names_df["Name"].apply(lambda x: process.extractOne(x, ids_df_names))

    names_df["Matched_Name"] = matches.apply(lambda x: x[0])
    names_df["Similarity_Score"] = matches.apply(lambda x: x[1])
    result_df = names_df.merge(ids_df, left_on="Matched_Name", right_on="Name", how="left")
    result_df = result_df[["Name_x", "Matched_Name", "Similarity_Score", "ID"]]
    result_df.columns = ["Original_Name", "Matched_Name", "Similarity_Score", "ID"]
    st.dataframe(result_df)

    # Allow users to specify the output file path
    user_path = st.text_input("Enter the output file path (e.g., /path/to/results.xlsx):")

    if st.button("Save File"):
        if user_path:
            # Ensure the file has a .xlsx extension
            if not user_path.endswith(".xlsx"):
                user_path += ".xlsx"

            # Ensure the directory exists
            directory = os.path.dirname(user_path)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)

            # Save the file
            try:
                result_df.to_excel(user_path, index=False)
                st.success(f"File saved successfully at {user_path}")
            except Exception as e:
                st.error(f"Error saving file: {e}")
        else:
            st.warning("Please enter a valid file path.")
