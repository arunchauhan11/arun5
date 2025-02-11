import streamlit as st
import pandas as pd

def save_to_excel(data, filename, sheet_name='Sheet1'):
    """
    Save data to an Excel file using pandas
    
    Parameters:
    data: Can be a dictionary, list of lists, or pandas DataFrame
    filename: String, path where the Excel file should be saved
    sheet_name: String, name of the sheet in Excel (default 'Sheet1')
    """
    # If data is a dictionary, convert to DataFrame
    if isinstance(data, dict):
        df = pd.DataFrame(data)
    # If data is a list of lists, convert to DataFrame
    elif isinstance(data, list):
        df = pd.DataFrame(data[1:], columns=data[0])
    # If data is already a DataFrame, use it directly
    elif isinstance(data, pd.DataFrame):
        df = data
    else:
        raise TypeError("Data must be a dictionary, list of lists, or DataFrame")
    
    # Save to Excel
    try:
        df.to_excel(filename, sheet_name=sheet_name, index=False)
        print(f"Data successfully saved to {filename}")
    except Exception as e:
        print(f"Error saving data: {str(e)}")

# Example usage with different data types
if __name__ == "__main__":
    # Example 1: Dictionary
    dict_data = {
        'Name': ['John', 'Alice', 'Bob'],
        'Age': [28, 24, 32],
        'City': ['New York', 'San Francisco', 'Chicago']
    }
    save_to_excel(dict_data, 'dictionary_data.xlsx')
    
    # Example 2: List of lists
    list_data = [
        ['Name', 'Score', 'Grade'],  # Headers
        ['John', 85, 'A'],
        ['Alice', 92, 'A'],
        ['Bob', 78, 'B']
    ]
    save_to_excel(list_data, 'list_data.xlsx')
    
    # Example 3: DataFrame
    df_data = pd.DataFrame({
        'Product': ['Laptop', 'Phone', 'Tablet'],
        'Price': [1200, 800, 400],
        'Stock': [15, 25, 10]
    })
    save_to_excel(df_data, 'dataframe_data.xlsx')
