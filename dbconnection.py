import pandas as pd

# Replace 'your_excel_file.xlsx' with the actual name of your Excel file
excel_file_path = r'Databaseviewer.xlsx'

# Read Excel file into a pandas DataFrame
try:
    df = pd.read_excel("Databaseviewer.xlsx")
    print("Connected to Excel file successfully.")
    # Now 'df' contains the data from your Excel file
    print(df.head())  # Display the first few rows of the DataFrame
except pd.errors.EmptyDataError:
    print("Excel file is empty or does not exist.")
except Exception as e:
    print(f"Error in connection: {e}")
