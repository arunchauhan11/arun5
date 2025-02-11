import streamlit as st
import pandas as pd
from io import BytesIO
import base64

def main():
    st.title("Excel Data Entry App")
    
    # Initialize session state for storing data
    if 'data' not in st.session_state:
        st.session_state.data = []
    
    # Create input form
    with st.form("data_form"):
        name = st.text_input("Name")
        age = st.number_input("Age", min_value=0, max_value=120)
        city = st.text_input("City")
        submitted = st.form_submit_button("Add Entry")
        
        if submitted:
            # Add new data to session state
            st.session_state.data.append({
                "Name": name,
                "Age": age,
                "City": city
            })
            st.success("Data added successfully!")
    
    # Display current data
    if st.session_state.data:
        st.subheader("Current Data")
        df = pd.DataFrame(st.session_state.data)
        st.dataframe(df)
        
        # Create Excel download button
        if st.button("Download Excel File"):
            # Create Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Data', index=False)
                
                # Get workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Data']
                
                # Add some formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#C9C9C9',
                    'border': 1
                })
                
                # Apply header formatting
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Auto-adjust columns' width
                for column in df:
                    column_length = max(df[column].astype(str).apply(len).max(),
                                     len(column))
                    col_idx = df.columns.get_loc(column)
                    work