import time
start_time = time.time()

import os
import copy
import pandas as pd
import numpy as np
import openpyxl
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
from PIL import Image
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from fuzzywuzzy import process  # Needed for fuzzy matching
import contextlib

# ---------------------------- PAGE CONFIG ----------------------------
st.set_page_config(page_title="Excel Consolidator", layout="wide", page_icon="📂")

st.markdown("""
    <style>
    #MainMenu, footer, header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ---------------------------- THEME TOGGLE ----------------------------
if 'dark_theme' not in st.session_state:
    st.session_state.dark_theme = False

def toggle_theme():
    st.session_state.dark_theme = not st.session_state.dark_theme

if st.session_state.dark_theme:
    st.markdown("<style>body { background-color: #1e1e1e; color: white; }</style>", unsafe_allow_html=True)
else:
    st.markdown("<style>body { background-color: white; color: black; }</style>", unsafe_allow_html=True)

# ---------------------------- SIDEBAR ----------------------------
logo = Image.open("ipec.jpg")
st.sidebar.image(logo, width=250)
st.sidebar.title("Settings ⚙️")
page = st.sidebar.selectbox("Select a page:", ["Home", "Excel Consolidator", "About"])

if st.sidebar.button("Toggle Dark/Light Theme"):
    toggle_theme()

# ---------------------------- PAGE LOGIC ----------------------------
if page == "Home":
    st.title("AML-CFT QRT MODEL 📊")
    st.markdown("Welcome to the AML-CFT QRT Automation System.")

elif page == "About":
    st.title("About This App")
    st.markdown("""
    Developed by the **Actuarial Team**.  
    Purpose: Combine multiple pension-related Excel files automatically into a single file for easy analysis.  
    **Contact:** actuarial@ipec.co.zw  
    **Version:** 1.0 (April 2025)
    """)

elif page == "Excel Consolidator":
    st.title("Excel File Consolidator 📂")
    st.markdown("Upload multiple Excel files (they must have a sheet called 'sheet1'):")
    
    
elif page == "🔍 Quantitative Analysis":
    st.title("📊 Quantitative Analysis Dashboard")
submenu = st.sidebar.radio("Choose Analysis Category:", [
"👤 Customers",
"📦 Products & Services",
"🧑‍💼 Use of Intermediaries",
"🌍 Geo Location of Clients",
"💳 Mode of Payment",
"🚨 Suspicious & Large Cash Transactions"
])

if submenu == "👤 Customers":
    st.header("👤 Customer Analysis")

    # File uploader
import streamlit as st
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from io import BytesIO
import copy
import os
import contextlib

# File uploader
uploaded_files = st.file_uploader("Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} files uploaded successfully!")

    combined_workbook = openpyxl.Workbook()
    combined_workbook.remove(combined_workbook.active)

    # Initialize intermediary data holder
    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(uploaded_file, data_only=False)
            filename = uploaded_file.name

            if "sheet1" in wb.sheetnames:
                sheet = wb["sheet1"]
                new_sheet = combined_workbook.create_sheet(title=filename[:-5])

                for row in sheet.iter_rows():
                    for cell in row:
                        new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            new_cell.font = copy.copy(cell.font)
                            new_cell.fill = copy.copy(cell.fill)
                            new_cell.border = copy.copy(cell.border)
                            new_cell.alignment = copy.copy(cell.alignment)
                            new_cell.number_format = cell.number_format

                for merged_range in sheet.merged_cells.ranges:
                    new_sheet.merge_cells(str(merged_range))

                st.success(f"✅ Successfully imported {filename} data")
            else:
                st.warning(f"⚠️ Skipping {uploaded_file.name} (No 'sheet1' found)")

        except InvalidFileException:
            st.error(f"❌ Skipping invalid file: {uploaded_file.name}")

    # -------------------- Download Combined File --------------------
    output = BytesIO()
    combined_workbook.save(output)
    output.seek(0)

    # Suppress output display
    with contextlib.redirect_stdout(open(os.devnull, 'w')):
        st.download_button(
            label="📥 Download Combined Workbook",
            data=output,
            file_name="combined_workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

            
  ###################################################################################################################################      
xls = pd.ExcelFile(output)
Inter_df = pd.DataFrame()
st.write("Use of Intermediaries")
group_names = ['brokers', 'agents', 'Direct clients', 'other specify']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "III. Use of Intermediaries (see note 3 below)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "III. Use of Intermediaries (see note 3 below)"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "IV. Geographic Location of Clients (as at the end of the reporting period)"].empty:
            end_row = df[df.iloc[:, 0] == "IV. Geographic Location of Clients (as at the end of the reporting period)"].index[0] - 1  # exclude the row with the end label
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'Nil' string with 0
            selected_data.replace('Nil', 0, inplace=True)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            Inter_df = pd.concat([Inter_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "use of intermediaries", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
Inter_df.columns = column_names + list(Inter_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
Inter_df = Inter_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    Inter_df[col] = pd.to_numeric(Inter_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
Inter_df[numeric_columns] = Inter_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 80  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
Inter_grouped_df = Inter_df[Inter_df['use of intermediaries'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
Inter_grouped_df = Inter_grouped_df.groupby('use of intermediaries')[numeric_columns].sum()


Inter_grouped_df = Inter_grouped_df.reset_index()


################################################################################################################################
xls = pd.ExcelFile(output)
geoloc_df = pd.DataFrame()

# Specify the group names that we want to filter by
group_names = ['a)    Specific High risk areas in Zimbabwe  (e.g. areas with small scale gold mining)(see note 4)', 'b)      Specific High risk foreign jurisdictions (List provided in glossary of terms)','      i.    UN/FATF Sanctioned countries', '      ii.    European Union Sanctioned countries','      iii.   Other jurisdictions considered high-risk']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "IV. Geographic Location of Clients (as at the end of the reporting period)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "IV. Geographic Location of Clients (as at the end of the reporting period)"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "VI.  Source of Funds"].empty:
            end_row = df[df.iloc[:, 0] == "VI.  Source of Funds"].index[0] - 1  # exclude the row with the end label
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the geoloc DataFrame
            geoloc_df = pd.concat([geoloc_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "Geographic location of Clients", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
geoloc_df.columns = column_names + list(geoloc_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
geoloc_df = geoloc_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    geoloc_df[col] = pd.to_numeric(geoloc_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
geoloc_df[numeric_columns] = geoloc_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 60  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
summary_geoloc_df = geoloc_df[geoloc_df['Geographic location of Clients'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
summary_geoloc_df = summary_geoloc_df.groupby('Geographic location of Clients')[numeric_columns].sum()



##########################################################################################################################################
products_combined_df = pd.DataFrame()
# Specify the group names that we want to filter by
group_names = [' Services ', ' Tombstone ',' Savings ', ' other ',' Annuities ', ' Guaranteed Education Plan ',' Whole Life ', ' Finance ',' Endowment ','  Pension', ' Education',' Credit',' Funeral ',' GLA ','Legal', ' Group Life',' Funds ','Ecosure']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "II. Products & Services Offered (as at the end of the reporting period)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "II. Products & Services Offered (as at the end of the reporting period)"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "III. Use of Intermediaries (see note 3 below)"].empty:
            end_row = df[df.iloc[:, 0] == "III. Use of Intermediaries (see note 3 below)"].index[0] - 1  # exclude the row with the end label
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            products_combined_df = pd.concat([products_combined_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "Products", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
products_combined_df.columns = column_names + list(products_combined_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
products_combined_df = products_combined_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    products_combined_df[col] = pd.to_numeric(products_combined_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
products_combined_df[numeric_columns] = products_combined_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 20  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
products_summary_df = products_combined_df[products_combined_df['Products'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
products_summary_df = products_summary_df.groupby('Products')[numeric_columns].sum().reset_index()
##################################################################################################################################
# Initialize an empty DataFrame
source_of_funds_consolidated_df = pd.DataFrame()

# Specify the group names that we want to filter by
group_names = ['salaried', 'self employed', 'formal trading business']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "VI.  Source of Funds"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "VI.  Source of Funds"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "V. Mode of Payment by Customer for Products and Services provided (year to date)"].empty:
            end_row = df[df.iloc[:, 0] == "V. Mode of Payment by Customer for Products and Services provided (year to date)"].index[0] - 1  # exclude the row with the end label
            
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            source_of_funds_consolidated_df = pd.concat([source_of_funds_consolidated_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "Source of funds", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
source_of_funds_consolidated_df.columns = column_names + list(source_of_funds_consolidated_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
source_of_funds_consolidated_df = source_of_funds_consolidated_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    source_of_funds_consolidated_df[col] = pd.to_numeric(source_of_funds_consolidated_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
source_of_funds_consolidated_df[numeric_columns] = source_of_funds_consolidated_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 60  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
source_of_funds_summary_df = source_of_funds_consolidated_df[source_of_funds_consolidated_df['Source of funds'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
source_of_funds_summary_df = source_of_funds_summary_df.groupby('Source of funds')[numeric_columns].sum()

#######################################################################################################################################
# Initialize an empty DataFrame
payment_mode_consolidated_df = pd.DataFrame()

# Specify the group names that we want to filter by
group_names = ['cash', 'Electronic transfer','Stop order']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "V. Mode of Payment by Customer for Products and Services provided (year to date)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "V. Mode of Payment by Customer for Products and Services provided (year to date)"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "VII. Suspicious and Large Cash Transaction Reports (year to date)"].empty:
            end_row = df[df.iloc[:, 0] == "VII. Suspicious and Large Cash Transaction Reports (year to date)"].index[0] - 1  # exclude the row with the end label
            
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            payment_mode_consolidated_df = pd.concat([payment_mode_consolidated_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "Mode of payment", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
payment_mode_consolidated_df.columns = column_names + list(payment_mode_consolidated_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
payment_mode_consolidated_df = payment_mode_consolidated_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    payment_mode_consolidated_df[col] = pd.to_numeric(payment_mode_consolidated_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
payment_mode_consolidated_df[numeric_columns] = payment_mode_consolidated_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 60  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
payment_mode_summary_df = payment_mode_consolidated_df[payment_mode_consolidated_df['Mode of payment'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
payment_mode_summary_df = payment_mode_summary_df.groupby('Mode of payment')[numeric_columns].sum()



#############################################################################################################################

# Initialize an empty DataFrame
cash_trans_consolidated_df = pd.DataFrame()

# Specify the group names that we want to filter by
group_names = ['large cash transactions reports filed', 'Suspicious Transaction Filed','Electronic Funds reports']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "VII. Suspicious and Large Cash Transaction Reports (year to date)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "VII. Suspicious and Large Cash Transaction Reports (year to date)"].index[0] + 1  # start from the next row

        if not df[df.iloc[:, 0] == "2. AML/CFT CONTROLS"].empty:
            end_row = df[df.iloc[:, 0] == "2. AML/CFT CONTROLS"].index[0] - 1  # exclude the row with the end label
            
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            cash_trans_consolidated_df = pd.concat([cash_trans_consolidated_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "Suspicious and Large Cash Transaction", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
cash_trans_consolidated_df.columns = column_names + list(cash_trans_consolidated_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
cash_trans_consolidated_df = cash_trans_consolidated_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    cash_trans_consolidated_df[col] = pd.to_numeric(cash_trans_consolidated_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
cash_trans_consolidated_df[numeric_columns] = cash_trans_consolidated_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 60  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
cash_trans_summary_df = cash_trans_consolidated_df[cash_trans_consolidated_df['Suspicious and Large Cash Transaction'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
cash_trans_summary_df = cash_trans_summary_df.groupby('Suspicious and Large Cash Transaction')[numeric_columns].sum()



####################################################################################################################################
customers_combined_df = pd.DataFrame()

# Specify the group names that we want to filter by
group_names = ['           i.     Domestic', '           ii.    Foreign', 'b.  Domestic Legal Clients ', 'e.  NGOs (Charities, Foundations etc)']

# Iterate through each sheet in the workbook
for sheet_name in xls.sheet_names:
    # Read the sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Check if the specified boundaries exist
    if not df[df.iloc[:, 0].str.strip() == "I. Customers (as at the end of the reporting period)"].empty:
        start_row = df[df.iloc[:, 0].str.strip() == "I. Customers (as at the end of the reporting period)"].index[0] + 2  # start from the next row

        if not df[df.iloc[:, 0] == "II. Products & Services Offered (as at the end of the reporting period)"].empty:
            end_row = df[df.iloc[:, 0] == "II. Products & Services Offered (as at the end of the reporting period)"].index[0] - 2  # exclude the row with the end label
            
            
            # Select the specified range of rows and all columns
            selected_data = df.iloc[start_row:end_row + 1, :].copy()
            # Add a new column with the sheet name
            selected_data.insert(0, 'Sheet Name', sheet_name)
            # Replace NaN with 0
            selected_data.fillna(0, inplace=True)
            # Replace 'nil' string with 0
            selected_data.replace('nil', 0, inplace=True)
            
            # Filter out rows where all values are zero (empty rows)
            selected_data = selected_data.loc[~(selected_data.iloc[:, 1:] == 0).all(axis=1)]

            # Append the selected data to the combined DataFrame
            customers_combined_df = pd.concat([customers_combined_df, selected_data], ignore_index=True)

# Add the desired column names to the first six columns only
column_names = ["company name", "customers", "number of clients ZWL", "total amount ZWL", 
                "number of clients USD", "total amount USD"]
customers_combined_df.columns = column_names + list(customers_combined_df.columns[len(column_names):])  # Keep extra columns intact

# Now, select only the first 6 columns to keep the data for only those columns
customers_combined_df = customers_combined_df.iloc[:, :6]

# Convert relevant numeric columns to numeric, coercing errors (non-numeric) to NaN
numeric_columns = ["number of clients ZWL", "total amount ZWL", "number of clients USD", "total amount USD"]
for col in numeric_columns:
    customers_combined_df[col] = pd.to_numeric(customers_combined_df[col], errors='coerce')

# Replace NaN values with 0 (optional, depending on how you want to handle missing values)
customers_combined_df[numeric_columns] = customers_combined_df[numeric_columns].fillna(0)

# Function to match strings with fuzzy matching
def fuzzy_match(row_value, group_names):
    best_match = process.extractOne(row_value, group_names)
    return best_match[1] >= 60  # Threshold of 80 for fuzzy matching score

# Apply fuzzy matching to filter rows based on 'use of intermediaries'
customers_summary_df = customers_combined_df[customers_combined_df['customers'].apply(lambda x: fuzzy_match(str(x), group_names))]

# Ensure numeric columns are summed and non-numeric columns are handled separately
customers_summary_df = customers_summary_df.groupby('customers')[numeric_columns].sum()




#############################################################################################################################################################
st.write("## Overview Dashboard 📊")
# Sort DataFrames
top_zwl_df = products_summary_df.sort_values("total amount ZWL", ascending=False).head(5)
top_usd_df = products_summary_df.sort_values("total amount USD", ascending=False).head(5)

# Create two columns
col1, col2 = st.columns(2)

# Pie chart for ZWL
with col1:
    fig1, ax1 = plt.subplots(figsize=(4, 4))  # Smaller figure
    wedges, texts, autotexts = ax1.pie(
        top_zwl_df["total amount ZWL"],
        labels=top_zwl_df["Products"],
        autopct="%1.0f%%",
        startangle=90,
        textprops={'fontsize': 8}  # Smaller font
    )
    ax1.axis("equal")
    plt.title("ZWL Business by Product", fontsize=10)
    st.pyplot(fig1)

# Pie chart for USD
with col2:
    fig2, ax2 = plt.subplots(figsize=(4, 4))  # Smaller figure
    wedges, texts, autotexts = ax2.pie(
        top_usd_df["total amount USD"],
        labels=top_usd_df["Products"],
        autopct="%1.0f%%",
        startangle=90,
        textprops={'fontsize': 8}  # Smaller font
    )
    ax2.axis("equal")
    plt.title("USD Business by Product", fontsize=10)
    st.pyplot(fig2)
##################################################################################################################################################################
st.write("# Use of Intermediaries 🤝")
st.write(Inter_df.head())
st.write(Inter_grouped_df)


st.write("# Geographic Location of Clients 🌍 (as at the end of the reporting period)")
st.write(geoloc_df)
st.write(summary_geoloc_df)

st.write("## Products & Services Offered 📦 (as at the end of the reporting period)")
st.write(products_summary_df)

st.write("## Source of Funds 💰")
st.write(source_of_funds_summary_df)


st.write("## Mode of Payment by Customer for Products and Services Provided 💳 (year to date)")
st.write(payment_mode_summary_df)


st.write("## Suspicious and Large Cash Transaction Reports 🚨 (year to date)")
st.write(cash_trans_summary_df)


st.write("## Inherent Risk Factors ⚠️")
st.write(customers_summary_df)
#################################################################################################################################################################
import pandas as pd
import streamlit as st
from io import BytesIO

# Your DataFrames dictionary
combined_summaries_workbook = {
    "customers_combined_df": customers_combined_df,
    "customers_summary_df": customers_summary_df,
    "cash_trans_consolidated_df": cash_trans_consolidated_df,
    "cash_trans_summary_df": cash_trans_summary_df,
    "payment_mode_consolidated_df": payment_mode_consolidated_df,
    "payment_mode_summary_df": payment_mode_summary_df,
    "source_of_funds_consolidated_df": source_of_funds_consolidated_df,
    "source_of_funds_summary_df": source_of_funds_summary_df,
    "products_combined_df": products_combined_df,
    "products_summary_df": products_summary_df,
    "summary_geoloc_df": summary_geoloc_df,
    "geoloc_df": geoloc_df,
    "Inter_df": Inter_df,
    "Inter_grouped_df": Inter_grouped_df
}

# Sheet names
sheet_names = {
    "customers_combined_df": "Customers",
    "customers_summary_df": "Cust_Sum",
    "cash_trans_consolidated_df": "Cash_trans",
    "cash_trans_summary_df": "Cash_trns_sum",
    "payment_mode_consolidated_df": "mode_of_payment",
    "payment_mode_summary_df": "Pay_Sum",
    "source_of_funds_consolidated_df": "source_of_funds",
    "source_of_funds_summary_df": "Funds_Sum",
    "products_combined_df": "products",
    "products_summary_df": "Prod_Sum",
    "summary_geoloc_df": "Geoloc_Sum",
    "geoloc_df": "Geoloc",
    "Inter_df": "Intermed",
    "Inter_grouped_df": "Intermed"
}

# Streamlit app
st.title("Download Summaries Excel File")

# Create the Excel file in memory
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    for name, df in combined_summaries_workbook.items():
        # Even if df is empty, create the sheet (optional: you can skip empty dfs if you prefer)
        if df is not None:
            df.to_excel(writer, sheet_name=sheet_names.get(name, name), index=False)
output.seek(0)  # Reset pointer to the beginning

# Download button
st.download_button(
    label="📥 Download Summaries Excel File",
    data=output,
    file_name="combined_summaries_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


##################################################################################################################################################################
import random
from docx import Document
from io import BytesIO
import streamlit as st

# Generate random data
companies_risk_profiled = round(random.random()*10, 1)
total_companies = round(random.random()*10, 1)
domestic_peps_percentage = round(random.random()*10, 1)
domestic_peps_amount = round(random.random()*10, 1)
domestic_peps_usd_percentage = round(random.random()*10, 1)
domestic_peps_usd_amount = round(random.random()*10, 1)
foreign_peps_reported = round(random.random()*10, 1)
trusts_percentage = round(random.random()*10, 1)
ngo_percentage = round(random.random()*10, 1)
high_net_worth_percentage_zig = round(random.random()*10, 1)
high_net_worth_percentage_usd = round(random.random()*10, 1)
agents_used = round(random.random()*10, 1)
brokers_used = round(random.random()*10, 1)
bancassurance_used = round(random.random()*10, 1)
direct_clients_used = round(random.random()*10, 1)
zig_business_agents_percentage = round(random.random()*10, 1)
usd_business_agents_percentage = round(random.random()*10, 1)
zig_direct_clients_percentage = round(random.random()*10, 1)

# Create Word Document
doc = Document()
doc.add_heading('ANTI-MONEY LAUNDERING AND COMBATING THE FINANCING OF TERRORISM AND PROLIFERATION FINANCING', level=1)
doc.add_paragraph("12 out of 12 life assurance companies submitted their third quarter 2024 AML/CFT/CPF returns.")
doc.add_heading('INHERENT RISK FACTORS AND CONTROLS', level=2)
doc.add_paragraph('The following observations were made from the submissions received during the quarter under review:')
doc.add_heading('Client Risk', level=2)
doc.add_paragraph(f"{companies_risk_profiled} out of {total_companies} entities risk-profiled their clients in Q3 2024.")
doc.add_paragraph(f"• Domestic PEPs: {domestic_peps_percentage}% (ZiG{domestic_peps_amount}), {domestic_peps_usd_percentage}% (US${domestic_peps_usd_amount}).")
doc.add_paragraph(f"• {foreign_peps_reported} foreign PEPs were reported.")
doc.add_paragraph(f"• Trusts: {trusts_percentage}%, NGOs: {ngo_percentage}%.")
doc.add_paragraph(f"• High net-worth individuals: {high_net_worth_percentage_zig}% (ZiG), {high_net_worth_percentage_usd}% (USD).")
doc.add_paragraph("The following chart shows the ZiG and USD business by client risk categories:")
doc.add_paragraph(f"The ZiG{domestic_peps_amount} and US${domestic_peps_usd_amount} amounts refer to high-risk clients only.")

doc.add_heading('Product Risk', level=2)
doc.add_paragraph('• Funeral products: 73%, Group Life: 15%, Conventional: 12%.')

doc.add_heading('Use of Intermediaries', level=1)
doc.add_paragraph(f"• Agents used: {agents_used}/12, Brokers: {brokers_used}, Bancassurance: {bancassurance_used}, Direct Clients: {direct_clients_used}")
doc.add_paragraph(f"• ZiG via Agents: {zig_business_agents_percentage}%, USD via Agents: {usd_business_agents_percentage}%")
doc.add_paragraph(f"• Direct Clients accounted for {zig_direct_clients_percentage}% of total business.")

# Save document
word_stream = BytesIO()
doc.save(word_stream)
word_stream.seek(0)

# Display feedback and button
st.success("Word report generated successfully!")
st.download_button(
    label="📄 Download Word Report",
    data=word_stream,
    file_name="AML_CFT_Report_Third_Quarter_2024.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


st.divider()
st.caption(f"App runtime: {round(time.time() - start_time, 2)} seconds.")
#################################################################################################################################################################













