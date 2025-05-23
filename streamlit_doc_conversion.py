# IMPORTS STATEMENTS
import streamlit as st
import pandas as pd
import numpy as np
import io
import csv


# PAYROLL METHODS
# Function to reformat the input data
def process_pr_data(input_file):
    # Read in the PR data
    raw_PR = pd.read_excel(input_file, sheet_name=1, header=None)
    #print(raw_PR.head())  # Print the first few rows to understand its structure
    print(raw_PR.shape)   # Check the number of rows and columns

    # Find the index where "DEPT" appears in column 7 (index 6)
    dept_row_index = raw_PR[raw_PR[6] == "DEPT"].index[0]

    # Subset rows from the 'DEPT' row onward, and select columns 7 to 9 (indices 5-8)
    PR_1 = raw_PR.iloc[dept_row_index:, 6:10]
    #print(PR_1.head)  # Print the first few rows to understand its structure
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

    # Print out column data types (for debugging)
    print(PR_4.dtypes)

    # Step 7: Check the file format
    required_columns = ['DEPT', 'ACCT', 'DEBITS', 'CREDITS']
    if not all(col in PR_4.columns for col in required_columns):
        raise ValueError("Input file must have 'Dept', 'Acct', 'Debit', and 'Credit' columns.")
    print("Passed required columns test. ")
    pr_data = PR_4

    return pr_data
# Function to generate the output data
def create_pr_file (processed_pr_data, journal_date, accounting_period, description_1):
    print("now entering output file creation")
    # Creating debit lines
    debit_lines = processed_pr_data[processed_pr_data['DEBITS'].notna()].copy()  # Filter rows where DEBITS is not NaN
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
    credit_lines = processed_pr_data[processed_pr_data['CREDITS'].notna()].copy()  # Filter rows where CREDITS is not NaN
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
    if st:
        pr_file = io.BytesIO()  # In-memory file for Streamlit (binary mode)
    else:
        pr_file = io.StringIO()  # In-memory file for terminal (text mode)
    gltrn_df.to_csv(pr_file, index=False, header=False, sep=",", quotechar='"', quoting=csv.QUOTE_ALL)
    pr_file.seek(0)  # Go to the beginning of the in-memory file

    return pr_file
# Payroll Main
def mainPR(uploaded_file):
    processed_data = process_pr_data(uploaded_file)
    print("data has been processed.")

    # Test the file creation function
    journal_date = "010125"
    accounting_period = "01"
    description_1 = "Payroll Entry"
    print("variables have been set.")
    pr_final = create_pr_file(processed_data, journal_date, accounting_period, description_1)

    # Print or save the output as needed
    print("File created successfully. Ready for download.")
    print(pr_final)

    output_content = pr_final.getvalue()
    print(output_content)  # This will print the file content to the terminal

    return pr_final

# CIGNA METHODS 
# Function to reformat the input data
def process_cig_data(input_file):
    # Read in the Cigna data
    raw_Cigna = pd.read_excel(input_file, sheet_name=3, header=None)
    print("Raw Cigna Data:")
    print(raw_Cigna.head())
    print(raw_Cigna.shape)

    # Set start and end rows (based on the "Employee ID" text in the first column)
    start_row = raw_Cigna[raw_Cigna[0] == "Employee ID"].index[0]
    end_row = raw_Cigna[raw_Cigna[0].isna() & (raw_Cigna.index > start_row)].index[0] - 1

    # Crop the data between start and end rows
    cropped_cig = raw_Cigna.iloc[start_row:end_row + 1, :]
    print("Cropped Cigna Data:")
    print(cropped_cig.head())
    print(cropped_cig.shape)

    # Set column names (first row in cropped data)
    cropped_cig.columns = cropped_cig.iloc[0]
    cropped_cig = cropped_cig.iloc[1:].reset_index(drop=True)  # Remove the first row
    print("Named Cropped Data:")
    print(cropped_cig.head())

    data_csv = cropped_cig

    # Delete unnecessary columns (adjust column names based on the actual data)
    data_csv = data_csv.drop(data_csv.columns[[2, 12]], axis=1)

    print("Modified Data:")
    print(data_csv.head())

    # Normalize Employee Name columns (to avoid issues with spaces/case differences)
    data_csv['Employee Name'] = data_csv['Employee Name'].str.strip().str.lower()
    #print(data_csv.columns)
    
    # Load the employee departments list and normalize Employee.Name column
    master_depts = pd.read_excel("employee_depts_master.xlsx")
    master_depts['Employee.Name'] = master_depts['Employee.Name'].str.strip().str.lower()
    master_depts.rename(columns={'Employee.Name': 'Employee Name'}, inplace=True)
    #print(master_depts.columns)

    # Merge the employee departments with the data (based on normalized Employee.Name)
    data_csv = data_csv.merge(master_depts, on="Employee Name", how="left")

    # Drop 'Subgrp ID' column and replace it with 'Dept.Acct'
    if 'Subgrp.ID' in data_csv.columns:
        data_csv.drop(columns=['Subgrp.ID'], inplace=True)
    # Now Dept.Acct is in the data_csv, it replaces Subgrp ID

    # Reorder columns (add 'Dept.Acct' at the correct place and make sure it's clean)
    data_csv = data_csv[['Employee Name', 'Dept.Acct'] + [col for col in data_csv.columns if col not in ['Employee Name', 'Dept.Acct']]]

    print("Departmental Data:")
    print(data_csv.head())

    cig_data = data_csv
    return cig_data  
# Function to generate the output data
def create_cig_file(processed_cig_data, journal_date, accounting_period, description_1, credit_acct):
    # Summarizing the data by Dept.Acct
    summary_data = processed_cig_data.groupby('Dept.Acct').agg(
        Sum_Medical=('Medical', 'sum'),
        Premium=('Amount Due (1)', 'sum'),
        Claims_Allocated=('Claims Funding (3)', 'sum'),
        Total=('Total (4)', 'sum')
    ).reset_index()
    
    # Step 1: Summing the necessary columns
    totals_row = summary_data[['Sum_Medical', 'Premium', 'Claims_Allocated', 'Total']].sum()

    # Step 2: Manually set 'Dept.Acct' to 'TOTAL'
    totals_row['Dept.Acct'] = 'TOTAL'

    # Step 3: Convert the totals_row (which is a Series) into a DataFrame (transpose)
    totals_row_df = totals_row.to_frame().T

    # Step 4: Append the totals row to the summary_data DataFrame
    summary_data = pd.concat([summary_data, totals_row_df], ignore_index=True)

    print("Summary Data:")
    print(summary_data)

    # Creating debit lines
    debit_lines = summary_data[summary_data['Dept.Acct'] != "TOTAL"].copy()
    debit_lines['Field1'] = "00000"
    debit_lines['Field2'] = "000100000" + accounting_period + "RE00000"
    debit_lines['Field3'] = "000"
    debit_lines['Field4'] = journal_date
    debit_lines['Field5'] = description_1
    debit_lines['Field6'] = ""
    debit_lines['Field7'] = debit_lines['Dept.Acct'].apply(lambda x: f"{x[:3]}00000{x[-4:]}")
    #debit_lines['Field8'] = debit_lines['Total'].apply(lambda x: f"{x*100:.0f}")  # Format total as integer
    debit_lines['Field8'] = debit_lines['Total'].apply(lambda x: f"{round(x, 2):.2f}")  # Round to 2 decimal places
    debit_lines['Field9'] = ""

    # Creating credit lines
    total_sum = summary_data[summary_data['Dept.Acct'] == "TOTAL"]['Total'].iloc[0]
    credit_line = pd.DataFrame({
        'Field1': ["00000"],
        'Field2': [f"000100000{accounting_period}RE00000"],
        'Field3': ["000"],
        'Field4': [journal_date],
        'Field5': [description_1],
        'Field6': [""],
        'Field7': [f"00000000{credit_acct}"],
        #'Field8': [f"{-total_sum * 100:.0f}"],  # Negative for credit
        'Field8': [f"{round(-total_sum, 2):.2f}"],  # Round to 2 decimal places and make it negative for credit
        'Field9': [""]
    })

    # Combine debit and credit lines
    gltrn_df = pd.concat([debit_lines, credit_line], ignore_index=True)
    gltrn_df = gltrn_df.filter(regex="^Field")  # Keep only the "Field" columns

    print("GLTRN Data:")
    print(gltrn_df)


    # Save the output to a .txt file
    if st:
        cig_file = io.BytesIO()  # In-memory file for Streamlit (binary mode)
    else:
        cig_file = io.StringIO()  # In-memory file for terminal (text mode)
    gltrn_df.to_csv(cig_file, index=False, header=False, sep=",", quotechar='"', quoting=csv.QUOTE_ALL)
    cig_file.seek(0)  # Go to the beginning of the in-memory file

    return cig_file
# Cigna Main
def mainCig(uploaded_file):
    processed_data = process_cig_data(uploaded_file)
    print("data has been processed.")

    # Test the file creation function
    journal_date = "010125"
    accounting_period = "01"
    description_1 = "Cigna Entry"
    credit_acct = "1130"
    print("variables have been set.")
    cig_final = create_cig_file(processed_data, journal_date, accounting_period, description_1, credit_acct)

    # Print or save the output as needed
    print("File created successfully. Ready for download.")
    print(cig_final)
    
    output_content = cig_final.getvalue()
    print(output_content)  # This will print the file content to the terminal


    return cig_final

# FIRST-TIME GIVERS METHODS  
# def process_ftg_data(input_file):
#     return file

# def create_ftg_file():
#     return file

# def mainFTG(uploaded_file):
#     return file


# TESTING (Outside Streamlit)
# if __name__ == "__main__":
#     # PAYROLL TEST
#     uploaded_PR = 'PR Journal Entry_03.25.2025-1.xlsx'  # Replace with the path to your test file
#     mainPR(uploaded_PR)

#     # CIGNA TEST
#     uploaded_Cig = 'GroupPremiumStatementRpt_03.2025.xlsx'  # Replace with the path to your test file
#     mainCig(uploaded_Cig)

#     # ARENA FTG TEST
#     uploaded_ftg = 'insert_name'  # Replace with the path to your test file
#     mainFTG(uploaded_ftg)


# STREAMLIT SETUP
st.title("File Import & Conversion Application")
# Step 1: Select the file type (Payroll or Cigna)
file_type = st.radio("Select the file type you're uploading:", ['Arena: First-Time Givers', 'Payroll', 'Cigna', ])

if file_type == 'Arena: First-Time Givers':
    st.header("Arena File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Upload an Arena-Downloaded Excel file", type="xlsx")
    #journal_date = st.text_input("Journal Date:", value="010125")
    #accounting_period = st.text_input("Accounting Period:", value="01")
    description = st.text_input("Description of Report", value="First-time Givers mm.yy")

elif file_type == 'Payroll':
    st.header("Payroll File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Choose an Excel file for Payroll", type="xlsx")
    journal_date = st.text_input("Journal Date:", value="010125")
    accounting_period = st.text_input("Accounting Period:", value="01")
    description_1 = st.text_input("Description for Journal Entry:", value="Payroll Entry xx.xx.xx")
    
    # Run the script when the button is pressed
    if st.button("Run Payroll Script"):
        if uploaded_file is not None:
            # Process the payroll data
            processed_data = process_pr_data(uploaded_file)
            
            # Create the output file
            output_file = create_pr_file(processed_data, journal_date, accounting_period, description_1)

            st.success("Payroll file processed and ready for download!")

            # Provide download button for the payroll output
            st.download_button(
                label="Download Payroll File",
                data=output_file,
                file_name="GLTRN2000.txt",
                mime="text/csv"
            )
        else:
            st.error("Please upload a payroll file.")

elif file_type == 'Cigna':
    st.header("Cigna File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Choose an Excel file for Cigna", type="xlsx")
    journal_date = st.text_input("Journal Date:", value="010125")
    accounting_period = st.text_input("Accounting Period:", value="01")
    description_1 = st.text_input("Description for Journal Entry:", value="Cigna Entry xx.xx.xx")
    credit_acct = "1130"
    
    # Run the script when the button is pressed
    if st.button("Run Cigna Script"):
        if uploaded_file is not None:
            # Process the Cigna data
            processed_data = process_cig_data(uploaded_file)
            
            # Create the output file
            output_file = create_cig_file(processed_data, journal_date, accounting_period, description_1, credit_acct)

            st.success("Cigna file processed and ready for download!")

            # Provide download button for the Cigna output
            st.download_button(
                label="Download Cigna File",
                data=output_file,
                file_name="GLTRN2000.txt",
                mime="text/csv"
            ) 
        else:
            st.error("Please upload a Cigna file.")
