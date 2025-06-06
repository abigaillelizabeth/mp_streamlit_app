# IMPORTS STATEMENTS
import streamlit as st
#import toml
import pandas as pd
import numpy as np
import csv
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import tempfile
import datetime
import io
import os
import zipfile


# ARENA METHODS  
# Function to reformat the input data
def process_arena_data(input_file):
    # Read in the PR data
    raw_arena = pd.read_excel(input_file, sheet_name=0, header=None)
    #print(raw_arena.head())  # Print the first few rows to understand its structure
    #print(raw_arena.shape)   # Check the number of rows and columns

    # remove extra columns
    raw_arena = raw_arena.iloc[:, list(range(0, 14)) + [15]]
    #print(raw_arena.head())

    # Group by Family Id (column 0) and Person ID (column 1)
    raw_arena.iloc[1:, 14] = raw_arena.iloc[1:, 14].astype(float)
    grouped_arena = raw_arena.groupby([0, 1], as_index=False).agg({
        2: 'first',    # Keep the first value for 'Last Name' (column 2)
        3: 'first',    # Keep the first value for 'First Name' (column 3)
        4: 'first',    # Keep the first value for 'Nick Name' (column 4)
        5: 'first',    # Keep the first value for 'Spouse Title' (column 5)
        6: 'first',    # Keep the first value for 'Spouse Last Name' (column 6)
        7: 'first',    # Keep the first value for 'Spouse First Name' (column 7)
        8: 'first',    # Keep the first value for 'Spouse Nick Name' (column 8)
        9: 'first',    # Keep the first value for 'Address' (column 9)
        10: 'first',   # Keep the first value for 'City' (column 10)
        11: 'first',   # Keep the first value for 'State' (column 11)
        12: 'first',   # Keep the first value for 'Zip' (column 12)
        13: 'first',   # Keep the first value for 'Email' (column 13)
        15: 'sum'      # Sum the 'Contribution Fund Amount' (column 15)
    })

    # Rename the columns for clarity
    grouped_arena.columns = ['Family Id', 'Person ID', 'Last Name', 'First Name', 'Nick Name', 
                             'Spouse Title', 'Spouse Last Name', 'Spouse First Name', 'Spouse Nick Name', 
                             'Address', 'City', 'State', 'Zip', 'Email', 'Total Contribution Fund Amount']

    # print("printing grouped arena")
    # print(grouped_arena.shape)
    # print(grouped_arena.head(20))

    # Add a blank column "Title" at index 3
    grouped_arena.insert(3, 'Title', '')  # Insert the new "Title" column at index 3 and leave it blank

    # Sort by Last name
    sorted_arena = grouped_arena.sort_values(by=['Last Name'])

    # TO_DO: if first name (index 4) == Nick Name (index 5) == blank, separate those columns to be at the very bottom of the output
    # Step 1: Filter rows where 'First Name' and 'Nick Name' are both blank
    blank_names = sorted_arena[(sorted_arena['First Name'].isna()) & (sorted_arena['Nick Name'].isna())]

    # Step 2: Filter out those rows from the main DataFrame
    sorted_arena = sorted_arena[~((sorted_arena['First Name'].isna()) & (sorted_arena['Nick Name'].isna()))]

    # Step 3: Append the rows with blank names at the bottom of the DataFrame
    sorted_arena = pd.concat([sorted_arena, blank_names])

    # Remove rows that are identical to the column name
    sorted_arena = sorted_arena[~(sorted_arena == sorted_arena.columns).all(axis=1)]

    arena_data = sorted_arena
    return arena_data
# Function to generate the output data
def create_arena_file(processed_arena_data, is_streamlit = True):
    if is_streamlit:
        # If running in Streamlit, keep the file in memory (no save to disk)
        arena_file = io.BytesIO()  # In-memory file for Streamlit (binary mode)
        processed_arena_data.to_excel(arena_file, index=False, engine='openpyxl')  # Use openpyxl for Excel output
        arena_file.seek(0)  # Go to the beginning of the in-memory file
    else:
        # If NOT running in Streamlit, save the file to disk
        output_file_path = "Sorted_Mailing_List.xlsx"
        processed_arena_data.to_excel(output_file_path, index=False, engine='openpyxl')  # Save to disk
        #processed_arena_data.to_excel(output_file_path, index=False)
        print(f"File successfully saved to {output_file_path}")
        arena_file = None  # Return None or any placeholder since file is saved on disk

    return arena_file
# Arena Main
def mainArena(uploaded_file):
    processed_data = process_arena_data(uploaded_file)
    print("data has been processed.")

    # Test the file creation function
    arena_final = create_arena_file(processed_data, is_streamlit = False)
    #print(arena_final)

    if arena_final != None:
        print("File created successfully. Ready for download.")
    else:
        print("File has been saved to disk.")

    return arena_final
# Function to run arena methods  
def runArenaMain():
    st.header("Arena File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Upload an Arena-Downloaded Excel file", type="xlsx")

    # Run the script when the button is pressed
    if st.button("Create Donor Summary"):
        if uploaded_file is not None:
            # Process the payroll data
            processed_data = process_arena_data(uploaded_file)
            
            # Create the output file
            output_file = create_arena_file(processed_data)

            st.success("Donor mailing list processed and ready for download!")

            # Provide download button for the payroll output
            st.download_button(
                label="Download Donor Summary",
                data=output_file,
                file_name="Sorted_Mailing_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                #mime="text/csv"
            )
        else:
            st.error("Please upload an arena file.")


# PAYROLL METHODS
# Function to reformat the input data
def process_pr_data(input_file):
    # Read in the PR data
    raw_PR = pd.read_excel(input_file, sheet_name=1, header=None)
    #print(raw_PR.head())  # Print the first few rows to understand its structure
    #print(raw_PR.shape)   # Check the number of rows and columns

    # Find the index where "DEPT" appears in column 7 (index 6)
    dept_row_index = raw_PR[raw_PR[6] == "DEPT"].index[0]

    # Subset rows from the 'DEPT' row onward, and select columns 7 to 9 (indices 5-8)
    PR_1 = raw_PR.iloc[dept_row_index:, 6:10]
    #print(PR_1.head)  # Print the first few rows to understand its structure
    #print(PR_1.shape)   # Check the number of rows and columns

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
    #print(PR_4.dtypes)

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
    pr_final = create_pr_file(processed_data, journal_date, accounting_period, description_1)

    # Print & save the output as needed
    print("File created successfully. Ready for download.")
    #print(pr_final)
    #output_content = pr_final.getvalue()
    #print(output_content)  # This will print the file content to the terminal

    return pr_final
# Function to run payroll methods  
def runPayroll():
    st.header("Payroll File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Choose an Excel file for Payroll", type="xlsx")
    journal_date = st.text_input("Journal Date:", value="010125")
    accounting_period = st.text_input("Accounting Period:", value="01")
    description_1 = st.text_input("Description for Journal Entry:", value="Payroll Entry xx.xx.xx")
    
    # Run the script when the button is pressed
    if st.button("Generate Payroll JE File"):
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


# CIGNA METHODS 
# Function to reformat the input data
def process_cig_data(input_file):
    # Read in the Cigna data
    raw_Cigna = pd.read_excel(input_file, sheet_name=3, header=None)
    #print("Raw Cigna Data:")
    #print(raw_Cigna.head())
    #print(raw_Cigna.shape)

    # Set start and end rows (based on the "Employee ID" text in the first column)
    start_row = raw_Cigna[raw_Cigna[0] == "Employee ID"].index[0]
    end_row = raw_Cigna[raw_Cigna[0].isna() & (raw_Cigna.index > start_row)].index[0] - 1

    # Crop the data between start and end rows
    cropped_cig = raw_Cigna.iloc[start_row:end_row + 1, :]
    #print("Cropped Cigna Data:")
    #print(cropped_cig.head())
    #print(cropped_cig.shape)

    # Set column names (first row in cropped data)
    cropped_cig.columns = cropped_cig.iloc[0]
    cropped_cig = cropped_cig.iloc[1:].reset_index(drop=True)  # Remove the first row
    #print("Named Cropped Data:")
    #print(cropped_cig.head())

    data_csv = cropped_cig

    # Delete unnecessary columns (adjust column names based on the actual data)
    data_csv = data_csv.drop(data_csv.columns[[2, 12]], axis=1)

    #print("Modified Data:")
    #print(data_csv.head())

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

    #print("Departmental Data:")
    #print(data_csv.head())

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
    #print("Summary Data:")
    #print(summary_data)

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
    #print("GLTRN Data:")
    #print(gltrn_df)

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
    cig_final = create_cig_file(processed_data, journal_date, accounting_period, description_1, credit_acct)

    # Print & save the output as needed
    print("File created successfully. Ready for download.")


    return cig_final
# Function to run cigna methods  
def runCigna():
    st.header("Cigna File Upload")
    # Input Information
    uploaded_file = st.file_uploader("Choose an Excel file for Cigna", type="xlsx")
    journal_date = st.text_input("Journal Date:", value="010125")
    accounting_period = st.text_input("Accounting Period:", value="01")
    description_1 = st.text_input("Description for Journal Entry:", value="Cigna Entry xx.xx.xx")
    credit_acct = "1130"
    
    # Run the script when the button is pressed
    if st.button("Generate Cigna JE File"):
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


# CONTRIBUTIONS- ARENA
def arena_master_included(uploaded_arena_files):
    #print("running arena_master_included")
    arena_data_list = []  # List to store the processed data from each file

    # Sort the files by name (to check length and determine if master or batch file)
    uploaded_arena_files.sort(key=lambda x: len(x.name))  # Sort by the file name length

    master_file_found = False
    combined_arena_data = pd.DataFrame()  # Initialize the combined dataframe

    # Process each uploaded Arena file
    for uploaded_arena_file in uploaded_arena_files:
        # Read in the EZT data based on type
        filename = uploaded_arena_file.name.lower()

        if filename.endswith('.csv'):
            raw_contrib_arena = pd.read_csv(uploaded_arena_file)
        elif filename.endswith('.xlsx'):
            raw_contrib_arena = pd.read_excel(uploaded_arena_file, sheet_name=0, header=None)
        else:
            st.warning(f"Unsupported file type: {uploaded_arena_file.name}")
            continue

        # Check if the file name has more than 5 characters (indicating it's a master file)
        if len(uploaded_arena_file.name.split('.')[0]) > 5:
            # Master file logic
            print(f"Master file found: {uploaded_arena_file.name}")
            master_file_found = True

            # Use only the first row as the header
            raw_contrib_arena.columns = raw_contrib_arena.iloc[0]
            raw_contrib_arena = raw_contrib_arena.drop(index=0)
            raw_contrib_arena.reset_index(drop=True, inplace=True)

            combined_arena_data = raw_contrib_arena  # Initialize combined_arena_data with the raw master file

        else:
             # Extract the batch number (first 5 digits of the file name)
            batch_number = uploaded_arena_file.name.split('.')[0]

            # Batch file logic (files with 5-digit names)
            print(f"Processing batch file: {uploaded_arena_file.name}")
            processed_batch_file = arena_all_new([uploaded_arena_file])  # Process batch files and return a DataFrame
            processed_batch_file["Batch #"] = batch_number  # Add the "Batch #" column to the batch file

            # Append to the list for combining later
            arena_data_list.append(processed_batch_file)
    
    # If there are batch files, combine them with the master file
    if arena_data_list:
        combined_arena_data = pd.concat([combined_arena_data] + arena_data_list, ignore_index=True)

    return combined_arena_data

def arena_all_new(uploaded_arena_files):
    #print("running arena_all_new")
    arena_data_list = []  # List to store the processed data from each file

    for idx, uploaded_arena_file in enumerate(uploaded_arena_files):
        filename = uploaded_arena_file.name.lower()
        # Read in the Arena data based on type
        if filename.endswith('.csv'):
            raw_contrib_arena = pd.read_csv(uploaded_arena_file)
        elif filename.endswith('.xlsx'):
            raw_contrib_arena = pd.read_excel(uploaded_arena_file, sheet_name=0, header=None)
        else:
            st.warning(f"Unsupported file type: {uploaded_arena_file.name}")
            continue

        #raw_contrib_arena = pd.read_excel(uploaded_arena_file, sheet_name=0, header=None)

        # Call the arena_col_names function to adjust the column headers
        raw_contrib_arena = arena_col_names(raw_contrib_arena)

        # Add a column for "Batch #" based on the file name (5-digit string)
        batch_number = uploaded_arena_file.name.split('.')[0]  # Get the 5-digit file name without extension
        raw_contrib_arena["Batch #"] = batch_number  # Add the "Batch #" column to the dataframe

        # Append the processed data to the list
        arena_data_list.append(raw_contrib_arena)
        
        # If it's the first file, initialize the combined dataframe
        if idx == 0:
            combined_arena_data = raw_contrib_arena
        else:
            # If it's not the first file, append it to the combined dataframe
            combined_arena_data = pd.concat([combined_arena_data, raw_contrib_arena], ignore_index=True)
        
    return combined_arena_data

def arena_col_names(raw_contrib_arena):
    #print("running arena_col_names")
    # Manually combine the first two rows into one header row, but only if both rows have values
    new_columns = [
        f"{col1} {col2}" if pd.notna(col1) and pd.notna(col2) else col1 if pd.notna(col1) else col2
        for col1, col2 in zip(raw_contrib_arena.iloc[0], raw_contrib_arena.iloc[1])
    ]
    # Assign the new concatenated columns as the header
    raw_contrib_arena.columns = new_columns
    raw_contrib_arena = raw_contrib_arena.drop([0, 1])  # Drop the first two rows (header rows)
    # Move the data from row 1 (index 1) to the top (index 0)
    raw_contrib_arena.reset_index(drop=True, inplace=True)

    return raw_contrib_arena

def arena_merge(uploaded_arena_files):
    #print("running arena_merge")
    # Check if all files have a 5-digit name (without the file extension)
    all_five_digits = True
    
    # Iterate through each file in the uploaded files
    for arena_file in uploaded_arena_files:

        print(f"filename: {arena_file.name}")
        print(f"type: {type(arena_file)}")
        arena_file.seek(0)
        print(f"first 10 bytes: {arena_file.read(10)}")
        arena_file.seek(0)  # reset after peeking

        file_name = arena_file.name.split('.')[0]  # Remove file extension to check the name length
        if len(file_name) != 5 or not file_name.isdigit():  # Check if the length is not 5 or it's not numeric
            all_five_digits = False
            break  # No need to check further, as we already know the files don't match the condition
    
    # Call appropriate function based on the file name length
    if all_five_digits:
        combined_arena_data = arena_all_new(uploaded_arena_files)
    else:
        combined_arena_data = arena_master_included(uploaded_arena_files)

    return combined_arena_data

def arena_excel(combined_arena_data):
    combined_arena_data["Contribution Date"] = pd.to_datetime(combined_arena_data["Contribution Date"], errors="coerce")
    print("testing arena formatting")

    # Save to a temporary Excel file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            combined_arena_data.to_excel(writer, index=False, sheet_name="Arena Contributions")

        # Open workbook to apply formatting
        wb = load_workbook(tmp.name)
        ws = wb.active

        # Grab header row to locate each column index
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

        for row in ws.iter_rows(min_row=2):
            # print("testing this fucker #1")
            for i, header in enumerate(headers):
                # print("testing this fucker #2")
                cell = row[i]
                header_lower = str(header).lower().strip()
                if header_lower in ["amount", "contribution amount", "total"]:
                    # print("testing this fucker #3")
                    cell.number_format = '"$"#,##0.00'
                elif "date" in header_lower:
                    cell.number_format = 'mm/dd/yy'
                    print("testing datetime without string check")
                    if isinstance(cell.value, datetime.datetime):  # make sure it’s not a string
                        print("testing this fucker with a string check")
                        cell.number_format = 'MM/DD/YYYY'
                else:
                    # print("testing this fucker #5")
                    cell.number_format = 'General'

        wb.save(tmp.name)

        # Stream it into memory for download
        output = io.BytesIO()
        with open(tmp.name, "rb") as f:
            output.write(f.read())
        output.seek(0)

    return output

def runArenaContributions():
    #st.write("Arena Batch Files Upload")
    # Upload Arena Batch Files (Allow multiple files)
    uploaded_arena_files = st.file_uploader("Arena Batch Files Upload", type="xlsx", accept_multiple_files=True)
    # Process the uploaded Arena files
    if st.button("Import Arena Batches"):
        if uploaded_arena_files:
            # Call the process_contrib_arena to handle the file processing and combine the files
            combined_arena_data = arena_merge(uploaded_arena_files)

            print("Combined Arena Data")
            #print(combined_arena_data.dtypes)

            # Store the combined Arena data in session state
            st.session_state.arena_data = combined_arena_data
            #st.session_state.arena_file_output = combined_arena_data
            st.success("Arena batches processed and combined successfully")

            if combined_arena_data is not None and not combined_arena_data.empty:
                output = arena_excel(combined_arena_data)

            #     st.download_button(
            #         label="Download Merged Arena Data",
            #         data=output,
            #         file_name="merged_arena_batches.xlsx",
            #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #     )
            # else:
            #     st.error("No valid data to write to Excel.")
        else:
            st.error("Please upload at least one Arena batch.")
        #arena_file_output = arena_excel(combined_arena_data)
        return 
        

# CONTRIBUTIONS- EASY-TITHE
def ezt_merge(uploaded_ezt_data):
    print("running ezt_merge")
    ezt_data_list = []  # List to store the processed data from each file
    
    # Iterate through each file in the uploaded files
    for idx, ezt_file in enumerate(uploaded_ezt_data): 
        # Readin the EZT data based on type
        filename = ezt_file.name.lower()
        if filename.endswith('.csv'):
            raw_contrib_ezt = pd.read_csv(ezt_file)
        elif filename.endswith('.xlsx'):
            raw_contrib_ezt = pd.read_excel(ezt_file, sheet_name=0, header=0)
        else:
            continue  # Skip unsupported formats

        raw_contrib_ezt = raw_contrib_ezt[raw_contrib_ezt["Date"].notna()] #  Remove summary rows
        ezt_data_list.append(raw_contrib_ezt) # Append the processed data to the list

        if idx == 0: # If it's the first file, initialize the combined dataframe
            combined_ezt_data = raw_contrib_ezt 
        else: # If it's not the first file, append it to the combined dataframe
            combined_ezt_data = pd.concat([combined_ezt_data, raw_contrib_ezt], ignore_index=True)
    return combined_ezt_data

def ezt_excel(combined_ezt_data):
    # Save to a temporary Excel file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            combined_ezt_data.to_excel(writer, index=False, sheet_name="EZT Contributions")

        # Open workbook to apply formatting
        wb = load_workbook(tmp.name)
        ws = wb.active

        # Get the header row (row 1)
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

        for row in ws.iter_rows(min_row=2):  # start from second row to skip header
            for i, header in enumerate(headers):
                cell = row[i]
                header_lower = str(header).lower().strip()
                if header_lower == "gross gift":
                    cell.number_format = '"$"#,##0.00'
                elif header_lower == "date":
                    cell.number_format = 'mm/dd/yy'
                else:
                    cell.number_format = 'General'

        wb.save(tmp.name)

        # Stream it into memory for download
        output = io.BytesIO()
        with open(tmp.name, "rb") as f:
            output.write(f.read())
        output.seek(0)

    return output

def runEZTContributions():
    #st.write("EasyTithe Batch File Upload")
    # Upload the EZT batch file
    uploaded_ezt_files = st.file_uploader(
        "EasyTithe Batch Files Upload",
        type=["xlsx", "csv"],
        accept_multiple_files=True
    )
    #Run the script when the button is clicked for EZT file
    if st.button("Import EasyTithe Batches"):
        if uploaded_ezt_files:
            # Call the ezt_contributions to handle the file processing and combine the files
            combined_ezt_data = ezt_merge(uploaded_ezt_files)
            print("Combined EZT data")

            # Store the combined EZT data in session state
            st.session_state.ezt_data = combined_ezt_data 
            st.success("EasyTithe batches processed and combined successfully")
        
            # if combined_ezt_data is not None and not combined_ezt_data.empty:
            #     output = ezt_excel(combined_ezt_data)
            #     st.download_button(
            #         label="Download Merged EZT Data",
            #         data=output,
            #         file_name="merged_EZT_batches.xlsx",
            #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #     )
            # else:
            #     st.error("No valid data to write to Excel.")
        else:
            st.error("Please upload at least one EasyTithe batch.")


# CONTRIBUTIONS- MATCHING BY ID
def reorder_merged_columns(merged, arena_df, ezt_df):
    arena_cols = list(arena_df.columns)
    ezt_cols = list(ezt_df.columns)

    arena_id_cols = ["Arena Transaction ID", "Arena Batch #"]
    ezt_id_cols = ["EZT Transaction ID", "EZT Batch ID"]

    arena_other_cols = [col for col in arena_cols if col not in ["Transaction Detail", "Batch #"]]
    ezt_other_cols = [col for col in ezt_cols if col not in ["Transaction Number", "Batch ID"]]

    arena_other_cols = [col for col in arena_other_cols if col not in arena_id_cols]
    ezt_other_cols = [col for col in ezt_other_cols if col not in ezt_id_cols]

    final_order = arena_id_cols + arena_other_cols + ezt_id_cols + ezt_other_cols
    return [col for col in final_order if col in merged.columns]

def matching_logic(arena_df, ezt_df):
    # Make copies to avoid changing original data
    arena_df = arena_df.copy()
    ezt_df = ezt_df.copy()

    # Rename key columns for clarity
    arena_df = arena_df.rename(columns={
        "Transaction Detail": "Arena Transaction ID",
        "Batch #": "Arena Batch #"
    })
    ezt_df = ezt_df.rename(columns={
        "Transaction Number": "EZT Transaction ID",
        "Batch ID": "EZT Batch ID"
    })
    
    #print("ARENA COLUMNS:", arena_df.columns.tolist())
    #print("EZT COLUMNS:", ezt_df.columns.tolist())


    # Merge on transaction ID with indicator
    merged = pd.merge(
        arena_df,
        ezt_df,
        left_on="Arena Transaction ID",
        right_on="EZT Transaction ID",
        how="outer",
        indicator=True
    )

    # Extract match categories
    match_by_transaction_id = merged[merged["_merge"] == "both"].copy()
    arena_only = merged[merged["_merge"] == "left_only"].copy()
    ezt_only = merged[merged["_merge"] == "right_only"].copy()

    # Add Match Type Indicator Column
    match_by_transaction_id["Match Type"] = "Matched by Transaction ID"
    arena_only["Match Type"] = "Unmatched (Arena only)"
    ezt_only["Match Type"] = "Unmatched (EZT only)"

    # Drop the merge indicator before output
    for df in [match_by_transaction_id, arena_only, ezt_only]:
        df.drop(columns=["_merge"], inplace=True)

    # Apply consistent column order
    ordered_cols = reorder_merged_columns(merged, arena_df, ezt_df) + ["Match Type"]
    match_by_transaction_id = match_by_transaction_id[ordered_cols]
    arena_only = arena_only[ordered_cols]
    ezt_only = ezt_only[ordered_cols]

    # Build categorized Excel with labeled sections
    return categorized_matches(
        match_by_id_df=match_by_transaction_id,
        match_by_donor_df=pd.DataFrame(),  # Placeholder for future donor matching
        unmatched_df=pd.concat([arena_only, ezt_only], ignore_index=True)
    )
    # return match_by_transaction_id, pd.DataFrame(), pd.concat([arena_only, ezt_only], ignore_index=True)

# def categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df):
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "Matched Contributions"
        
#     def write_section(header, df, start_row):
#         # Write title row, bold, only col A populated
#         for col in range(1, ws.max_column + 1 or len(df.columns) + 1):
#             cell = ws.cell(row=start_row, column=col)
#             if col == 1:
#                 cell.value = header
#             else:
#                 cell.value = ""
#             cell.font = Font(bold=True)

#         # Write DataFrame, bold header row
#         for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row + 1):
#             for c_idx, value in enumerate(row, start=1):
#                 cell = ws.cell(row=r_idx, column=c_idx, value=value)
#                 #if r_idx == start_row + 1:
#                 #   cell.font = Font(bold=True)

#         return r_idx + 2  # Advance cursor

#     row_cursor = 1
#     row_cursor = write_section("Matched by Transaction ID", match_by_id_df, row_cursor)
#     if not match_by_donor_df.empty:
#         row_cursor = write_section("Matched by Donor Info", match_by_donor_df, row_cursor)
#     row_cursor = write_section("Unmatched Transactions", unmatched_df, row_cursor)

#     wb.save(tmp.name)
#     with open(tmp.name, "rb") as f:
#         final_output = f.read()

#     return final_output
def categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb = Workbook()
        ws = wb.active
        ws.title = "Matched Contributions"
        
        def write_section(header, df, start_row):
            # KEY FIX: Use DataFrame column count instead of ws.max_column
            num_columns = max(len(df.columns), 1)  # Ensure at least 1 column
            
            # Write title row - ONLY populate column A
            for col_idx in range(1, num_columns + 1):
                cell = ws.cell(row=start_row, column=col_idx)
                if col_idx == 1:
                    cell.value = header
                    #cell.font = Font(bold=True)
                else:
                    # Explicitly clear other cells in header row
                    cell.value = None

            # Write DataFrame content
            header_row = start_row + 1
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=header_row):
                for c_idx, value in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx == header_row:  # Bold column headers
                        cell.font = Font(bold=True)
            
            return r_idx + 2  # Add blank row between sections

        row_cursor = 1
        row_cursor = write_section("Matched by Transaction ID", match_by_id_df, row_cursor)
        
        if not match_by_donor_df.empty:
            row_cursor = write_section("Matched by Donor Info", match_by_donor_df, row_cursor)
            
        row_cursor = write_section("Unmatched Transactions", unmatched_df, row_cursor)

        wb.save(tmp.name)
        with open(tmp.name, "rb") as f:
            final_output = f.read()

    return final_output
      
# CONTRIBUTIONS - EXPORTING
def export_combined_excel(arena_df, ezt_df):
    # Format and return as .xlsx binary output
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
            # Write Arena data
            arena_df.to_excel(writer, index=False, sheet_name="Arena Contributions")
            # Write EZT data
            ezt_df.to_excel(writer, index=False, sheet_name="EZT Contributions")

        # Now format both sheets
        wb = load_workbook(tmp.name)

        for sheet_name, df in {
            "Arena Contributions": arena_df,
            "EZT Contributions": ezt_df
        }.items():
            ws = wb[sheet_name]
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            for row in ws.iter_rows(min_row=2):
                for i, header in enumerate(headers):
                    cell = row[i]
                    header_lower = str(header).lower().strip()
                    if header_lower in ["gross gift", "amount", "contribution amount", "total"]:
                        cell.number_format = '"$"#,##0.00'
                    elif "date" in header_lower:
                        cell.number_format = "mm/dd/yy"
                    else:
                        cell.number_format = "General"

        wb.save(tmp.name)

        # now convert to correct output file type
        output = io.BytesIO()
        with open(tmp.name, "rb") as f:
            output.write(f.read())
        output.seek(0)

    return output

def export_matched_excel(arena_df, ezt_df):
    match_by_id_df, match_by_donor_df, unmatched_df = matching_logic(arena_df, ezt_df)
    return categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df)

def export_full_report(arena_df, ezt_df, matched_data):

    match_by_id_df, match_by_donor_df, unmatched_df = matching_logic(arena_df, ezt_df)

    #arena_wb = load_workbook(matched_data)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            categorized_df = categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df)
            pd.read_excel(io.BytesIO(categorized_df)).to_excel(writer, sheet_name="Matched Contributions", index=False)
            arena_df.to_excel(writer, index=False, sheet_name="Arena Contributions")
            # for sheet in arena_wb.sheetnames:
            #     sheet_data = arena_wb[sheet]
            #     temp_df = pd.DataFrame(sheet_data.values)
            #     temp_df.columns = temp_df.iloc[0]
            #     temp_df = temp_df[1:]
            #     temp_df.to_excel(writer, sheet_name=sheet, index=False)
            ezt_df.to_excel(writer, index=False, sheet_name="EZT Contributions")

        with open(tmp.name, "rb") as f:
            output = io.BytesIO(f.read())
    output.seek(0)
    return output

# def export_full_report(arena_df, ezt_df): # CHAT CODE
#     # Get output from arena formatting function
#     arena_excel_output = arena_excel(arena_df)
#     arena_wb = load_workbook(io.BytesIO(arena_excel_output))
#     arena_ws = arena_wb.active

#     # Run matching logic
#     match_by_id_df, match_by_donor_df, unmatched_df = matching_logic(arena_df, ezt_df)

#     # Start a new workbook
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
#         wb = Workbook()
#         writer = pd.ExcelWriter(tmp.name, engine='openpyxl')
#         writer.book = wb

#         # --- Sheet 1: Matched Contributions (with section headers)
#         matched_sheet_bytes = categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df)
#         matched_wb = load_workbook(io.BytesIO(matched_sheet_bytes))
#         matched_ws = matched_wb.active
#         matched_new_ws = wb.create_sheet("Matched Contributions")

#         for row in matched_ws.iter_rows(values_only=True):
#             matched_new_ws.append(row)

#         # --- Sheet 2: Arena Contributions (formatted by arena_excel)
#         arena_new_ws = wb.create_sheet("Arena Contributions")
#         for row in arena_ws.iter_rows(values_only=True):
#             arena_new_ws.append(row)

#         # --- Sheet 3: EZT Contributions
#         ezt_df.to_excel(writer, index=False, sheet_name="EZT Contributions")

#         # Delete default "Sheet" if still present
#         if "Sheet" in wb.sheetnames:
#             std = wb["Sheet"]
#             wb.remove(std)

#         # Save
#         wb.save(tmp.name)

#         # Stream back to memory
#         output = io.BytesIO()
#         with open(tmp.name, "rb") as f:
#             output.write(f.read())
#         output.seek(0)

#     return output

def runMatchingFunctions():
    # set session state booleans for reference
    arena_ready = 'arena_data' in st.session_state
    ezt_ready = 'ezt_data' in st.session_state

    if arena_ready and ezt_ready:
        st.header("Contribution Report Download")
        # st.write("Choose how you'd like to export the contribution data:")

        # # Button 1 — Merged .xlsx workbook - 2 sheets
        # merged_excel = export_combined_excel(
        #     st.session_state.arena_data,
        #     st.session_state.ezt_data
        # )
        # st.download_button(
        #     label="Combined Arena & EZT Batches (.xlsx)",
        #     data=merged_excel,
        #     file_name="merged_contributions.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )

        # # Button 2 — Matched CSV file
        # matched_excel = export_matched_excel(
        #     st.session_state.arena_data,
        #     st.session_state.ezt_data
        # )
        # st.download_button(
        #     label="Matched Sheet Only (.xlsx)",
        #     data=matched_excel,
        #     file_name="matched_contributions.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )

        # Button 3 — Matched AND Merged Infomation - 3 sheets
        master_excel = export_full_report(
            st.session_state.arena_data,
            st.session_state.ezt_data, 
            # st.session_state.matched_data
        )
        st.download_button(
            label="Master Workbook (.xlsx)",
            data=master_excel,
            file_name="master_contributions_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.info("Upload both Arena and EZT files to access combined export options.")


# def categorized_matches(match_by_id_df, match_by_donor_df, unmatched_df):
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
#         # Start a writer
#         with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
#             # Create empty DataFrame just to initialize the sheet
#             pd.DataFrame().to_excel(writer, sheet_name="Matched Contributions", index=False)

#         # Open with openpyxl
#         wb = load_workbook(tmp.name)
#         ws = wb["Matched Contributions"]

#         def write_section(header, df, start_row):
#             ws.cell(row=start_row, column=1, value=header).font = Font(bold=True)
#             for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row + 1):
#                 for c_idx, value in enumerate(row, start=1):
#                     ws.cell(row=r_idx, column=c_idx, value=value)
#             return r_idx + 2  # Add spacing

#         row_cursor = 1
#         row_cursor = write_section("Matched by Transaction ID", match_by_id_df, row_cursor)
#         row_cursor = write_section("Matched by Donor Info", match_by_donor_df, row_cursor)
#         row_cursor = write_section("Unmatched Transactions", unmatched_df, row_cursor)

#         wb.save(tmp.name)

#         # Load back into memory
#         with open(tmp.name, "rb") as f:
#             final_output = f.read()

#     return final_output

# def match_data_logic(arena_df, ezt_df):
#     # Make copies to avoid changing original data
#     arena_df = arena_df.copy()
#     ezt_df = ezt_df.copy()

#     # Rename transaction columns for clarity
#     arena_df = arena_df.rename(columns={
#         "Transaction Detail": "Arena Transaction ID",
#         "Batch #": "Arena Batch #"
#     })
#     ezt_df = ezt_df.rename(columns={
#         "Transaction Number": "EZT Transaction ID",
#         "Batch ID": "EZT Batch ID"
#     })

#     # Merge on transaction IDs with an indicator to track match status
#     merged = pd.merge(
#         arena_df,
#         ezt_df,
#         left_on="Arena Transaction ID",
#         right_on="EZT Transaction ID",
#         how="outer",
#         indicator=True
#     )

#     # Reorder: transaction IDs first, then everything else
#     arena_id_cols = ["Arena Transaction ID",  "Arena Batch #"]
#     arena_other_cols = [col for col in arena_df.columns if col not in arena_id_cols + ["_merge"]]
#     ezt_id_cols = ["EZT Transaction ID", "EZT Batch ID"]
#     ezt_other_cols = [col for col in ezt_df.columns if col not in ezt_id_cols + ["_merge"]]
#     merged = merged[arena_id_cols + arena_other_cols + ezt_id_cols + ezt_other_cols + ["_merge"]]
    

#     # Sort rows: matched first, then arena-only, then ezt-only
#     # matched = merged[merged["_merge"] == "both"]
#     # arena_only = merged[merged["_merge"] == "left_only"]
#     # ezt_only = merged[merged["_merge"] == "right_only"]

#     # Combine all into one DataFrame
#     # final_df = pd.concat([matched, ezt_only, arena_only], ignore_index=True).drop(columns=["_merge"])


#     # Sorting for categorization
#     match_by_transaction_id = merged[merged["_merge"] == "both"].copy()
#     arena_only = merged[merged["_merge"] == "left_only"].copy()
#     ezt_only = merged[merged["_merge"] == "right_only"].copy()

#     # Adding the Matched column
#     match_by_transaction_id["Match Type"] = "Matched by Transaction ID"
#     arena_only["Match Type"] = "Unmatched (Arena only)"
#     ezt_only["Match Type"] = "Unmatched (EZT only)"

#     for df in [match_by_transaction_id, arena_only, ezt_only]:
#         df.drop(columns=["_merge"], inplace=True)
    
#     final_df = categorized_matches(
#         match_by_id_df=match_by_transaction_id,
#         match_by_donor_df=pd.DataFrame(),  # Placeholder until we build donor match
#         unmatched_df=pd.concat([arena_only, ezt_only], ignore_index=True)
# )
#     return final_df

# def export_matched_excel(arena_df, ezt_df):
#     # Merge the datasets
#     merged_df = match_data_logic(arena_df, ezt_df)

#     # Save to a temporary Excel file
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
#         with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
#             merged_df.to_excel(writer, index=False, sheet_name="Matched Contributions")

#         # Open workbook to apply formatting
#         wb = load_workbook(tmp.name)
#         ws = wb.active

#         # Get the header row
#         headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

#         # Format specific columns
#         for row in ws.iter_rows(min_row=2):
#             for i, header in enumerate(headers):
#                 cell = row[i]
#                 header_lower = str(header).lower().strip()

#                 if "amount" in header_lower or "gift" in header_lower:
#                     cell.number_format = '"$"#,##0.00'
#                 elif "date" in header_lower:
#                     cell.number_format = 'mm/dd/yy'
#                 else:
#                     cell.number_format = 'General'

#         wb.save(tmp.name)

#         # Stream to memory
#         output = io.BytesIO()
#         with open(tmp.name, "rb") as f:
#             output.write(f.read())
#         output.seek(0)

#     return output

# def export_full_report(arena_df, ezt_df):
#     matched_df = match_data_logic(arena_df, ezt_df)

#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
#         with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
#             matched_df.to_excel(writer, index=False, sheet_name="Matched Contributions")
#             arena_df.to_excel(writer, index=False, sheet_name="Arena Contributions")
#             ezt_df.to_excel(writer, index=False, sheet_name="EZT Contributions")

#         wb = load_workbook(tmp.name)
#         wb.save(tmp.name)

#         output = io.BytesIO()
#         with open(tmp.name, "rb") as f:
#             output.write(f.read())
#         output.seek(0)

#     return output


# Function to run all contribution methods  
def runContributions():
    #print("running runContributions")
    st.header("Contribution Reports Processing")
    # Upload Arena Batch Files (Allow multiple files)
    runArenaContributions()
    # Upload EZT batch file
    runEZTContributions()
    # Match the Contributions
    runMatchingFunctions()
    #print("contribution processing complete")


# STREAMLIT METHODS
# Authenticate user
def authenticate(username, password):
    # Loop through the credentials to match the username and password
    for user in st.secrets["credentials"]["user"]:
        if user["username"] == username and user["password"] == password:
            return True  # Successful authentication
    return False  # Failed authentication

# Call and implement the file methods
def call_methods():
    # When the user is logged in, show the rest of the app
    st.write("You are successfully logged in.")
    st.title("File Import & Conversion App")
    file_type = st.radio("Select the file type you want to convert:", 
                         ['Contribution Reports', 'Arena Mailing List', 'Payroll Workbook', 'Cigna Download'])
    # Run the correct set of methods based on user-selected file type
    if file_type == 'Contribution Reports':
        runContributions()
    elif file_type == 'Payroll Workbook':
        runPayroll()
    elif file_type == 'Cigna Download':
        runCigna()
    elif file_type == 'Arena Mailing List':
        runArenaMain()

# Set streamlit logic 
def run_gui():
    # Set up session state to track login status
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    # Show login form if not logged in
    if not st.session_state.logged_in:
        st.title("MP File Conversion Login Page")
        st.write("Please enter your username and password.")
        
        # Create input fields for username and password
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        # Login button to check credentials
        login_button = st.button("Login")

        if login_button:
        # Check if the credentials are valid
            if authenticate(username, password):
                # Successful login
                st.session_state.logged_in = True  # Store login status
                st.success("Login successful!")
                st.rerun()  # This is to refresh the page and hide the login form
            else:
                # Invalid credentials
                st.error("Invalid credentials. Please try again.")
    
    else:
        # Add a logout button
        if st.button("Logout"):
            st.session_state.logged_in = False  # Reset login status
            st.rerun()  # Refresh the app after logging out
        # Run the application
        call_methods()


# Streamit WITHOUT AUTH
if __name__ == "__main__":
    #print("running streamlit app")
    call_methods()
    
# Streamit WITH AUTH
#if __name__ == "__main__":
#    run_gui()


# #  CONTRIBUTIONS TERMINAL TESTING -- NOT RUNNING 
# def load_multiple_excels(folder_path):
#     combined_df = pd.DataFrame()
#     for filename in os.listdir(folder_path):
#         if filename.endswith(".xlsx") or filename.endswith(".xls"):
#             file_path = os.path.join(folder_path, filename)
#             df = pd.read_excel(file_path, header=None)  # read without headers
#             df = parse_column_headers(df)  # apply custom header formatting
#             combined_df = pd.concat([combined_df, df], ignore_index=True)
#     return combined_df

# def parse_column_headers(df):
#     new_columns = df.iloc[0].fillna('') + ' ' + df.iloc[1].fillna('')
#     df.columns = new_columns.str.strip()
#     return df.iloc[2:].reset_index(drop=True)

# if __name__ == "__main__":
#     arena_data = load_multiple_excels("arena_tests")
#     ezt_data = load_multiple_excels("ezt_tests")

#     result = export_full_report(arena_data, ezt_data)

#     with open("test_output.xlsx", "wb") as f:
#         f.write(result)

# # GENERAL TERMINAL TESTING
# if __name__ == "__main__":
    # # PAYROLL
    # uploaded_PR = 'PR Journal Entry_03.25.2025-1.xlsx'  # Replace with the path to your test file
    # mainPR(uploaded_PR)
    # # CIGNA
    # uploaded_Cig = 'GroupPremiumStatementRpt_03.2025.xlsx'  # Replace with the path to your test file
    # mainCig(uploaded_Cig)
    # # ARENA 
    # uploaded_arena = 'Arena Masterfile Tester.xlsx'  # Replace with the path to your test file
    # mainArena(uploaded_arena)

