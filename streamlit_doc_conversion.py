# IMPORTS STATEMENTS
import streamlit as st
import toml
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
    #print(cig_final)
    #output_content = cig_final.getvalue()
    #print(output_content)  # This will print the file content to the terminal

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
def runArena():
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


# CONTRIBUTION METHODS  
# Function to read in the Arena data
def process_contrib_arena(input_file):
    # Read in the Arena data
    raw_contrib_arena = pd.read_excel(input_file, sheet_name=0, header=None)

    print(raw_contrib_arena.head())  # Print the first few rows to understand its structure
    print(raw_contrib_arena.shape)   # Check the number of rows and columns

    return raw_contrib_arena
# Function to read in the EasyTithe data
def process_contrib_ezt(input_file):
    # Read in the EZT data
    raw_contrib_ezt = pd.read_excel(input_file, sheet_name=0, header=None)

    print(raw_contrib_ezt.head())  # Print the first few rows to understand its structure
    print(raw_contrib_ezt.shape)   # Check the number of rows and columns

    return raw_contrib_ezt
# Function to Match contribution data imports
def combine_contributions(arena_data, ezt_data):
    # Assuming the two dataframes have some common columns (e.g., "Family Id" or "Person ID")
    
    # Here we simply merge the two datasets on a common column
    # Adjust the merging strategy as needed (e.g., use a left join, inner join, etc.)
    combined_data = pd.merge(arena_data, ezt_data, on="Family Id", how="outer")
    
    # You can add any logic to clean the data or adjust the structure if needed

    return combined_data
# Function to run contribution methods  
def runContributions():
    st.header("Contribution Reports Processing")

    # Upload Arena Batch Files (Allow multiple files)
    st.write("Arena Batch Files Upload")
    uploaded_arena_files = st.file_uploader("Choose Arena batch files", type="xlsx", accept_multiple_files=True)
    # Process the uploaded Arena files
    if st.button("Import Arena Batches"):
        if uploaded_arena_files:
            arena_data_list = []
            for uploaded_arena_file in uploaded_arena_files:
                # Process each uploaded Arena file and append the result to a list
                arena_data = process_contrib_arena(uploaded_arena_file)
                arena_data_list.append(arena_data)

            # Combine all Arena data into one DataFrame
            combined_arena_data = pd.concat(arena_data_list, ignore_index=True)
            st.session_state.arena_data = combined_arena_data  # Store the combined Arena data in session state
            st.success("Arena batches processed and combined successfully")
        else:
            st.error("Please upload at least one Arena batch.")

    # Upload the EZT batch file
    st.write("EasyTithe Batch File Upload")
    uploaded_ezt_file = st.file_uploader("Choose an EasyTithe batch file", type="xlsx")
    # Run the script when the button is clicked for EZT file
    if st.button("Import EasyTithe Batch"):
        if uploaded_ezt_file is not None:
            # Process the EZT batch data
            ezt_data = process_contrib_ezt(uploaded_ezt_file)
            st.session_state.ezt_data = ezt_data  # Save the EZT data in session state
            st.success("EasyTithe batch processed")
        else:
            st.error("Please upload an EasyTithe batch.")

    # Combine the two files if both are uploaded
    if 'arena_data' in st.session_state and 'ezt_data' in st.session_state:
        if st.button("Combine Contribution Files"):
            combined_data = combine_contributions(st.session_state.arena_data, st.session_state.ezt_data)
            # Provide download button for the combined file
            st.download_button(
                label="Download Combined Contributions",
                data=combined_data.to_csv(index=False).encode(),  # Converting to CSV format
                file_name="combined_contribution_report.csv",
                mime="text/csv"
            )
        
    print("to be implemented")


# STREAMLIT METHODS
# Function to authenticate user
def authenticate(username, password):
    # Loop through the credentials to match the username and password
    for user in st.secrets["credentials"]["user"]:
        if user["username"] == username and user["password"] == password:
            return True  # Successful authentication
    return False  # Failed authentication

# Function to call and implement the file methods
def call_methods():
    # When the user is logged in, show the rest of the app
    st.write("You are successfully logged in.")
    st.title("File Import & Conversion App")
    file_type = st.radio("Select the file type you want to convert:", 
                         ['Arena Mailing List', 'Payroll Workbook', 'Cigna Download', 'Contribution Reports'])
    # Run the correct set of methods based on user-selected file type
    if file_type == 'Arena Mailing List':
        runArena()
    elif file_type == 'Payroll Workbook':
        runPayroll()
    elif file_type == 'Cigna Download':
        runCigna()
    elif file_type == 'Contribution Reports':
        runContributions()

# Function to set streamlit logic
def app():
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
        

# STREAMLIT TESTING
if __name__ == "__main__":
    app()

# # TERMINAL TESTING
# if __name__ == "__main__":
#     # PAYROLL
#     uploaded_PR = 'PR Journal Entry_03.25.2025-1.xlsx'  # Replace with the path to your test file
#     mainPR(uploaded_PR)
#     # CIGNA
#     uploaded_Cig = 'GroupPremiumStatementRpt_03.2025.xlsx'  # Replace with the path to your test file
#     mainCig(uploaded_Cig)
#     # ARENA 
#     uploaded_arena = 'Arena Masterfile Tester.xlsx'  # Replace with the path to your test file
#     mainArena(uploaded_arena)

