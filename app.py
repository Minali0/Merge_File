import pandas as pd
import streamlit as st

# Streamlit page title
st.title("Excel File Merger")

# File uploader with support for multiple file uploads
uploaded_files = st.file_uploader("Upload at least two Excel files", accept_multiple_files=True, type=["xlsx"])

# Ensure that at least two files are uploaded
if uploaded_files and len(uploaded_files) < 2:
    st.warning("Please upload at least two files to proceed.")

elif uploaded_files and len(uploaded_files) >= 2:
    # Initialize DataFrame with the first file
    merged_df = pd.read_excel(uploaded_files[0])
    merged_df = merged_df.rename(columns={"Amount": "Amount_file1"})

    # Sequentially merge other files
    for i, file in enumerate(uploaded_files[1:], start=2):
        df = pd.read_excel(file)
        df = df.rename(columns={"Amount": f"Amount_file{i}"})
        merged_df = pd.merge(merged_df, df, on='Invoice_no', how='outer')

    # Fill NaN values and calculate differences
    amount_columns = [col for col in merged_df.columns if 'Amount' in col]
    for col in amount_columns:
        merged_df[col] = merged_df[col].fillna(0)
    merged_df['amount_difference'] = merged_df[amount_columns[0]]
    for col in amount_columns[1:]:
        merged_df['amount_difference'] -= merged_df[col]

    # Define match status
    def get_match_status(row):
        status = []
        for i, col in enumerate(amount_columns, start=1):
            if row[col] == 0:
                status.append(f"not found in file_{i}")
        if not status and row['amount_difference'] == 0:
            return "matched"
        elif not status:
            return "unmatched"
        return ", ".join(status)

    merged_df['match_status'] = merged_df.apply(get_match_status, axis=1)

    # Display merged dataframe in Streamlit
    st.write("Merged Data Preview:")
    st.dataframe(merged_df)

    # Save to Excel and provide a download button
    output_file = "merged_invoices.xlsx"
    merged_df.to_excel(output_file, index=False)

    # Create a download button
    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Merged Excel File",
            data=file,
            file_name="merged_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
