# IMPORT STATEMENTS
import streamlit as st
import pandas as pd
import numpy as np
import io
import csv


# HELPER METHODS
# Function to reformat the input data
def process_data(input_file):
    # Read in the data
    raw_PR = pd.read_excel(input_file, sheet_name=1, header=None)
    print(raw_PR.head())  # Print the first few rows to understand its structure
    print(raw_PR.shape)   # Check the number of rows and columns

    # Find the index where "DEPT" appears in column 7 (index 6)
    dept_row_index = raw_PR[raw_PR[6] == "DEPT"].index[0]

    # Subset rows from the 'DEPT' row onward, and select columns 7 to 9 (indices 5-8)
    PR_1 = raw_PR.iloc[dept_row_index:, 6:10]

    print(PR_1.tail(20))  # Print the first few rows to understand its structure
    print(PR_1.shape)   # Check the number of rows and columns

    # Set column names
    PR_1.columns = PR_1.iloc[0]
    PR_1.columns = PR_1.columns.str.strip()  # Remove whitespace
    PR_2 = PR_1  # Create PR_2

    # Drop rows with NaN values in 'Dept' and 'Acct'
    PR_3 = PR_2.dropna(subset=[PR_2.columns[0], PR_2.columns[1]], how='all')

    # Drop rows with NaN values in 'Debit' and 'Credit'
    PR_4 = PR_3.dropna(subset=[PR_3.columns[2], PR_3.columns[3]], how='all')

    # Step 5: Convert to numericals
    #PR_4['DEBITS'] = pd.to_numeric(PR_4['DEBITS'], errors='coerce')  # Convert DEBITS to numeric, coercing errors to NaN
    #PR_4['CREDITS'] = pd.to_numeric(PR_4['CREDITS'], errors='coerce')  # Convert CREDITS to numeric, coercing errors to NaN

    PR_4.loc[:, 'DEBITS'] = pd.to_numeric(PR_4['DEBITS'], errors='coerce')
    PR_4.loc[:, 'CREDITS'] = pd.to_numeric(PR_4['CREDITS'], errors='coerce')


    # Print out column data types before conversion (for debugging)
    #print("Data types before conversion:")
    print(PR_4.dtypes)

    # Step 6: Convert all numeric columns to regular Python types (int or float)
    #for column in PR_4.select_dtypes(include=[np.int64, np.float64]).columns:
    #    PR_4[column] = PR_4[column].apply(lambda x: int(x) if isinstance(x, np.int64) else float(x))

    #PR_4['DEBITS'] = PR_4['DEBITS'].apply(lambda x: float(x) if pd.notnull(x) else 0)
    #PR_4['CREDITS'] = PR_4['CREDITS'].apply(lambda x: float(x) if pd.notnull(x) else 0.0)

    #for column in PR_4.select_dtypes(include=[np.int64, np.float64]).columns:
    #    PR_4[column] = PR_4[column].astype(object)  # Convert to Python-native types

    # Print out column data types after conversion (for debugging)
    #print("Data types after conversion:")
    #print(PR_4.dtypes)

    # Step 7: Check the file format
    required_columns = ['DEPT', 'ACCT', 'DEBITS', 'CREDITS']
    if not all(col in PR_4.columns for col in required_columns):
        raise ValueError("Input file must have 'Dept', 'Acct', 'Debit', and 'Credit' columns.")
    print("Passed required columns test. ")
    return PR_4




# def process_data(input_file):
#     # Read in the data
#     raw_PR = pd.read_excel(input_file, sheet_name=1, header=None)

#     # Check the columns to see what we're working with
#     print("Columns in the file:", raw_PR.columns)

#     # Step 1: Subset rows and columns
#     PR_1 = raw_PR.loc[raw_PR[6] == "DEPT", 6:9]

#     # Step 2: Changing column names
#     PR_1.columns = PR_1.iloc[0]  # Set column names to the first row
#     PR_1.columns = PR_1.columns.str.strip() # remove the whitespace
#     PR_2 = PR_1 # create PR_2

#     # Check the updated columns to ensure we have the expected ones
#     print("Columns after renaming:", PR_2.columns)

#     # Step 3: Remove rows with NaN in 'Dept' and 'Acct'
#     PR_3 = PR_2.dropna(subset=[PR_2.columns[0], PR_2.columns[1]], how='all')

#     # Step 4: Remove rows with NaN in 'Debit' and 'Credit'
#     PR_4 = PR_3.dropna(subset=[PR_3.columns[2], PR_3.columns[3]], how='all')

#     # Step 5: Convert to numericals
#     PR_4['DEBITS'] = pd.to_numeric(PR_4['DEBITS'], errors='coerce')  # Convert DEBITS to numeric, coercing errors to NaN
#     PR_4['CREDITS'] = pd.to_numeric(PR_4['CREDITS'], errors='coerce')  # Convert CREDITS to numeric, coercing errors to NaN

#     # Step 6: Convert all int64 columns to regular int
#     PR_4 = PR_4.apply(lambda col: col.astype(int) if col.dtype == 'int64' else col)

#     # Step 7: Check the file format
#     required_columns = ['DEPT', 'ACCT', 'DEBITS', 'CREDITS']
#     if not all(col in PR_4.columns for col in required_columns):
#         raise ValueError("Input file must have 'Dept', 'Acct', 'Debit', and 'Credit' columns.")
    
#     return PR_4




# Function to generate the output data
def file_creation(PR_4, journal_date, accounting_period, description_1):
    print("now entering output file creation")
    # Creating debit lines
    debit_lines = PR_4[PR_4['DEBITS'].notna()].copy()  # Filter rows where DEBITS is not NaN
    debit_lines['Field1'] = "00000"
    debit_lines['Field2'] = "000100000" + accounting_period + "JE00000"
    debit_lines['Field3'] = "000"
    debit_lines['Field4'] = journal_date
    debit_lines['Field5'] = description_1
    debit_lines['Field6'] = ""
    debit_lines['Field7'] = np.where(debit_lines['DEPT'].astype(str) == "0",
                                  "00000000" + debit_lines['ACCT'].astype(str),
                                  "0" + debit_lines['DEPT'].astype(str) + "00000" + debit_lines['ACCT'].astype(str))
    debit_lines['Field8'] = debit_lines['DEBITS'].apply(lambda x: f"{x:.2f}")  # Format to 2 decimal places
    debit_lines['Field9'] = ""
    #print(debit_lines.dtypes)

    # Creating credit lines
    credit_lines = PR_4[PR_4['CREDITS'].notna()].copy()  # Filter rows where CREDITS is not NaN
    credit_lines['Field1'] = "00000"
    credit_lines['Field2'] = "000100000" + accounting_period + "JE00000"
    credit_lines['Field3'] = "000"
    credit_lines['Field4'] = journal_date
    credit_lines['Field5'] = description_1
    credit_lines['Field6'] = ""
    credit_lines['Field7'] = np.where(credit_lines['DEPT'].astype(str) == "0", 
                                  "00000000" + credit_lines['ACCT'].astype(str),
                                  "0" + credit_lines['DEPT'].astype(str) + "00000" + credit_lines['ACCT'].astype(str))
    credit_lines['Field8'] = credit_lines['CREDITS'].apply(lambda x: f"{x * -1:.2f}" if x != 0 else "0.00")
    credit_lines['Field9'] = ""
    #print(credit_lines.dtypes)

    # Binding all lines (combining debit and credit lines)
    gltrn_df = pd.concat([debit_lines, credit_lines], ignore_index=True)
    gltrn_df = gltrn_df.filter(regex="^Field")  # Select only the columns that start with "Field"

    # Save the output to a .txt file
    output = io.StringIO()  # In-memory file- terminal CLI
    output = io.BytesIO() # In-memory file- streamlit
    gltrn_df.to_csv(output, index=False, header=False, sep=",", quotechar='"', quoting=csv.QUOTE_ALL)
    output.seek(0)  # Go to the beginning of the in-memory file

    return output


# # TESTING (Outside Streamlit)
# if __name__ == "__main__":
#     # Test the process_data function with a sample Excel file
#     uploaded_file = 'PR Journal Entry_03.25.2025-1.xlsx'  # Replace with the path to your test file
#     processed_data = process_data(uploaded_file)
#     print("data has been processed.")

#     # Test the file creation function
#     journal_date = "010125"
#     accounting_period = "01"
#     description_1 = "Payroll Entry"
#     print("variables have been set.")
#     output_file = file_creation(processed_data, journal_date, accounting_period, description_1)

#     # Print or save the output as needed
#     print("File created successfully. Ready for download.")
#     print(output_file)

#     output_content = output_file.getvalue()
#     print(output_content)  # This will print the file content to the terminal




# STREAMLIT SETUP
st.title("Payroll GLTRN File Export")
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
journal_date = st.text_input("Journal Date:", value="010125")
accounting_period = st.text_input("Accounting Period:", value="01")
description_1 = st.text_input("Description for Journal Entry:", value="Payroll Entry xx.xx.xx")


# BUTTON FUNCTIONALITY
if st.button("Process File"):
    if uploaded_file is not None:
        # Process the data
        processed_data = process_data(uploaded_file)
        
        # Display processed data
        json_data = processed_data.to_json(orient="split")
        #st.write(json_data)

        #st.write(processed_data)

        # Create the file and generate a download link
        output_file = file_creation(processed_data, journal_date, accounting_period, description_1)

        st.success("File processed and ready for download!")

        # Provide download button
        st.download_button(
            label="Download GLTRN",
            data=output_file,
            file_name="GLTRN2000.txt",
            mime="text/csv"
        )
        
        
    else:
        st.error("Please upload a file.")
