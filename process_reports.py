import pandas as pd
import re
import sqlite3
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# --- Version Number ---
__version__ = "2.3"

# --- Chart of Accounts Mapping ---
CHART_OF_ACCOUNTS = {
    'Gross Pay': {'acct': '65060', 'desc': 'Salaries & Wages : Employees'},
    'Commission': {'acct': '65062', 'desc': 'Salaries & Wages : Commission to Employees'},
    'Car Allowance': {'acct': '61000', 'desc': 'Automobile Expenses'},
    'ER_FICA': {'acct': '66502', 'desc': 'Payroll Taxes : FICA'},
    'FUTA': {'acct': '66503', 'desc': 'Payroll Taxes : FUTA'},
    'SUTA': {'acct': '66505', 'desc': 'Payroll Taxes : SUTA - PR'},
    'EE_FICA': {'acct': '23002', 'desc': 'Payroll Liabilities : FICA/FWH'},
    'State_WH': {'acct': '23004', 'desc': 'Payroll Liabilities : State W/H - PR'},
    'SDI': {'acct': '23001', 'desc': 'Payroll Liabilities : Disability - PR'},
    'Net_Pay': {'acct': '11090', 'desc': 'Payroll Exchange'}
}

# --- Department Mapping ---
DEPARTMENT_MATCH_LIST = [
    ('Montanez Ocasio', 'Warehouse'), ('Rosario Ramos', 'Warehouse'),
    ('Torres Ocasio', 'Warehouse'), ('Chimelis Crespo', 'Service Department'),
    ('Santiago Santiago', 'Service Department'), ('Albino Perez', 'Sales'),
    ('Aragon Rodriguez', 'Warehouse'), ('Palermo, Walter', 'Warehouse'),
    ('Rosario Cornejo', 'Accounting & Finance'), ('Jonathan Kieran Layton', 'Administration'),
    ('Silvia Z Layton', 'Administration'), ('Joel Pineda', 'Purchasing'),
    ('James Francisco Layton', 'Purchasing'), ('Jonathan Preston Layton', 'Purchasing'),
    ('Jorge A Ruiz', 'Warehouse'), ('Migdalia Sanchez', 'Warehouse')
]

def select_excel_file(title):
    """Opens a file dialog for the user to select an Excel file."""
    root = tk.Tk(); root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=[("Excel Files", "*.xlsx *.xls")])

def save_to_database(df, table_name, conn):
    """Saves a DataFrame to the SQLite database."""
    try:
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        print(f"\nSuccessfully saved data to table '{table_name}' in the database.")
    except Exception as e:
        print(f"\nError saving to database: {e}")

def display_database_tables(conn):
    """Connects to the DB and prints all tables."""
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    if not tables: print("\nDatabase is empty."); return
    print("\n" + "="*50); print("    DISPLAYING SAVED DATA"); print("="*50)
    for table_name in tables:
        table_name = table_name[0]
        print(f"\n--- Contents of table: {table_name} ---")
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
            print(df.to_string(index=False))
        except Exception as e:
            print(f"Could not read table {table_name}. Error: {e}")

def process_payroll_register(df):
    """Processes Payroll Register sheet data using a flexible department map."""
    payroll_data = {}
    current_employee_name = None
    for index, row in df.iterrows():
        first_cell_str = str(row.iloc[0])
        if pd.notna(row.iloc[0]) and "Associate ID:" in first_cell_str:
            current_employee_name = ' '.join(first_cell_str.split('\n')[0].strip().split())
            if current_employee_name not in payroll_data:
                department = 'Unknown'
                for key, dept_name in DEPARTMENT_MATCH_LIST:
                    if key in current_employee_name:
                        department = dept_name
                        break
                payroll_data[current_employee_name] = {'Department': department, 'Gross Pay': 0, 'Car Allowance': 0, 'Commission': 0}
        if current_employee_name and pd.notna(row.iloc[8]):
            gross_pay = 0.0
            try: gross_pay = float(str(row.iloc[8]).replace(',', ''))
            except (ValueError, TypeError): continue
            if pd.notna(row.iloc[6]) and "COM" in str(row.iloc[6]):
                payroll_data[current_employee_name]['Commission'] += gross_pay
            else:
                payroll_data[current_employee_name]['Gross Pay'] += gross_pay
            if len(row) > 11 and pd.notna(row.iloc[11]):
                match = re.search(r"CAL CarAllowance\s+\(([\d.]+)\)", str(row.iloc[11]))
                if match: payroll_data[current_employee_name]['Car Allowance'] += float(match.group(1))
        if "Dept. Total" in first_cell_str: current_employee_name = None
    summary_list = [{'Employee Name': name, **data} for name, data in payroll_data.items()]
    return pd.DataFrame(summary_list)

def process_statistical_summary(df):
    """Processes Statistical Summary sheet data."""
    data_to_process = []
    start_processing = False
    for index, row in df.iterrows():
        marker_val, description_val = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
        if description_val == "Total Taxes Debited": break
        if not start_processing and marker_val == "Taxes Debited": start_processing = True
        if start_processing and pd.notna(description_val) and description_val != 'nan':
            data_to_process.append({'Description': description_val, 'Value': row.iloc[2]})
    if not data_to_process: return pd.DataFrame()
    result_df = pd.DataFrame(data_to_process)
    result_df['Numeric Value'] = pd.to_numeric(result_df['Value'], errors='coerce').fillna(0)
    return result_df[result_df['Numeric Value'] > 0].copy()[['Description', 'Numeric Value']]

def create_journal_entry(conn, payroll_date):
    """Pulls data from DB, builds a detailed JE, and saves it to a CSV."""
    print(f"\nBuilding Journal Entry for date: {payroll_date}")
    try:
        register_df = pd.read_sql_query(f"SELECT * FROM payroll_register WHERE PayrollDate = '{payroll_date}'", conn)
        taxes_df = pd.read_sql_query(f"SELECT * FROM statistical_summary_taxes WHERE PayrollDate = '{payroll_date}'", conn)
        if register_df.empty or taxes_df.empty:
            print("No data found in the database for the specified date."); return

        tax_map = {
            'EE_FICA': taxes_df[taxes_df['Description'].str.contains('Social Security - EE|Medicare - EE', regex=True)]['Numeric Value'].sum(),
            'ER_FICA': taxes_df[taxes_df['Description'].str.contains('Social Security - ER|Medicare - ER', regex=True)]['Numeric Value'].sum(),
            'FUTA': taxes_df[taxes_df['Description'] == 'Federal Unemployment Tax']['Numeric Value'].sum(),
            'SUTA': taxes_df[taxes_df['Description'] == 'State Unemployment/Disability Ins - ER']['Numeric Value'].sum(),
            'State_WH': taxes_df[taxes_df['Description'] == 'State Income Tax']['Numeric Value'].sum(),
            'SDI': taxes_df[taxes_df['Description'] == 'State Disability Insurance - EE']['Numeric Value'].sum()
        }
        
        journal_lines = []
        entity = 2  # Assuming entity ID is 2 for the subsidiary
        for index, employee in register_df.iterrows():
            emp_name, emp_dept = employee['Employee Name'], employee['Department']
            if employee['Gross Pay'] > 0:
                acct = CHART_OF_ACCOUNTS['Gross Pay']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Gross Pay: {emp_name}", 'Department': emp_dept, 'Debit': employee['Gross Pay'], 'Credit': 0, 'Subsidiary': entity})
            if employee['Commission'] > 0:
                acct = CHART_OF_ACCOUNTS['Commission']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Commission: {emp_name}", 'Department': emp_dept, 'Debit': employee['Commission'], 'Credit': 0, 'Subsidiary': entity})
            if employee['Car Allowance'] > 0:
                acct = CHART_OF_ACCOUNTS['Car Allowance']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Car Allowance: {emp_name}", 'Department': emp_dept, 'Debit': employee['Car Allowance'], 'Credit': 0, 'Subsidiary': entity})

        # Aggregated lines
        acct = CHART_OF_ACCOUNTS['ER_FICA']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer FICA', 'Department': '', 'Debit': tax_map['ER_FICA'], 'Credit': 0, 'Subsidiary': entity})
        acct = CHART_OF_ACCOUNTS['FUTA']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer FUTA', 'Department': '', 'Debit': tax_map['FUTA'], 'Credit': 0, 'Subsidiary': entity})
        acct = CHART_OF_ACCOUNTS['SUTA']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer SUTA', 'Department': '', 'Debit': tax_map['SUTA'], 'Credit': 0, 'Subsidiary': entity})
        acct = CHART_OF_ACCOUNTS['EE_FICA']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employee FICA Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['EE_FICA'], 'Subsidiary': entity})
        acct = CHART_OF_ACCOUNTS['State_WH']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'State Income Tax Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['State_WH'], 'Subsidiary': entity})
        acct = CHART_OF_ACCOUNTS['SDI']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employee SDI Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['SDI'], 'Subsidiary': entity})

        je_df_temp = pd.DataFrame(journal_lines); net_pay = je_df_temp['Debit'].sum() - je_df_temp['Credit'].sum()
        acct = CHART_OF_ACCOUNTS['Net_Pay']; journal_lines.append({'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Payroll Cash Clearing for {payroll_date}", 'Department': '', 'Debit': 0, 'Credit': net_pay, 'Subsidiary': entity})

        final_je_df = pd.DataFrame(journal_lines)
        final_je_df['Date'] = payroll_date
        final_je_df[['Debit', 'Credit']] = final_je_df[['Debit', 'Credit']].round(2)
        
        # --- NEW: Balance Check ---
        total_debits = final_je_df['Debit'].sum()
        total_credits = final_je_df['Credit'].sum()

        if abs(total_debits - total_credits) > 0.01:
            print("\n" + "="*50); print("    !!! WARNING: Journal Entry is out of balance !!!"); print("="*50)
            print(f"Total Debits: ${total_debits:,.2f}"); print(f"Total Credits: ${total_credits:,.2f}")
            print("CSV file will not be created."); return

        print("\nSuccess! Debits and Credits are balanced: ${:,.2f}".format(total_debits))
        print("\n--- Generated Journal Entry ---")
        
        column_order = ['Date', 'Account', 'Memo', 'Department', 'Debit', 'Credit', 'Subsidiary']
        final_je_df = final_je_df[column_order]
        print(final_je_df.to_string(index=False))
        
        output_filename = input("\nEnter a filename to save this journal entry (e.g., JE_Export.csv): ")
        if output_filename:
            # --- NEW: Blank out zeros for the CSV file ---
            csv_df = final_je_df.copy()
            csv_df['Debit'] = csv_df['Debit'].apply(lambda x: '' if x == 0 else x)
            csv_df['Credit'] = csv_df['Credit'].apply(lambda x: '' if x == 0 else x)
            if not output_filename.lower().endswith('.csv'): output_filename += '.csv'
            csv_df.to_csv(output_filename, index=False)
            print(f"Journal Entry successfully saved to '{output_filename}'")

    except Exception as e:
        print(f"An error occurred while creating the journal entry: {e}")

def import_and_process_files():
    """Main function for the import workflow."""
    while True:
        date_str = input("\nPlease enter the Payroll Date for this batch (e.g., 07/15/2025): ")
        try:
            date_obj = datetime.strptime(date_str, '%m/%d/%Y'); payroll_date = date_obj.strftime('%m/%d/%Y')
            print(f"Using date: {payroll_date}"); break
        except ValueError: print("Invalid date format. Please use MM/DD/YYYY.")
    
    print("\n--- Step 1 of 2: Processing Payroll Register ---")
    file_path_register = select_excel_file("Select the Payroll Register Excel File")
    if file_path_register: process_and_save(file_path_register, "register", "payroll_register", payroll_date)
    else: print("File selection cancelled."); return

    print("\n--- Step 2 of 2: Processing Statistical Summary ---")
    file_path_summary = select_excel_file("Select the Statistical Summary Excel File")
    if file_path_summary: process_and_save(file_path_summary, "summary", "statistical_summary_taxes", payroll_date)
    else: print("File selection cancelled."); return

def process_and_save(file_path, processor_type, table_name, payroll_date):
    """Helper function to handle file processing and saving."""
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        print("\nWhich tab would you like to work on?")
        for i, name in enumerate(sheet_names): print(f"  {i+1}: {name}")
        choice_index = int(input("Enter the number for the tab: ")) - 1
        if not (0 <= choice_index < len(sheet_names)): print("Invalid selection."); return
        chosen_sheet = sheet_names[choice_index]
        print(f"\nProcessing tab: '{chosen_sheet}'...")
        df = pd.read_excel(file_path, sheet_name=chosen_sheet, header=None)
        
        result_df, display_df = None, None
        if processor_type == "register":
            result_df = process_payroll_register(df)
            display_df = result_df.copy()
            cols_to_display = ['Department', 'Employee Name', 'Gross Pay', 'Car Allowance', 'Commission']
            display_df = display_df[cols_to_display]
            for col in ['Gross Pay', 'Car Allowance', 'Commission']: display_df[col] = display_df[col].map('${:,.2f}'.format)
        elif processor_type == "summary":
            result_df = process_statistical_summary(df)
            display_df = result_df.copy()
            display_df.columns = ['Description', 'Amount']
            display_df['Amount'] = display_df['Amount'].map('{:,.2f}'.format)
        
        if result_df is not None and not result_df.empty:
            print("\n--- Processed Data ---"); print(display_df.to_string(index=False))
            save_choice = input("\nSave this table to the database? (y/n): ").lower()
            if save_choice == 'y':
                result_df['PayrollDate'] = payroll_date
                conn = sqlite3.connect('payroll_database.db')
                save_to_database(result_df, table_name, conn)
                conn.close()
    except Exception as e: print(f"An error occurred: {e}")

# --- Main Program Loop ---
if __name__ == "__main__":
    while True:
        print("\n" + "="*20 + " MAIN MENU " + "="*20)
        print(f" Payroll Report Processor v{__version__}\n")
        print("  1: Import and Process Files")
        print("  2: Display Saved Data from Database")
        print("  3: Create Journal Entry")
        print("  4: Quit")
        
        choice = input("\nEnter your choice (1-4): ")
        if choice == '1': import_and_process_files()
        elif choice == '2':
            try:
                conn = sqlite3.connect('payroll_database.db'); display_database_tables(conn); conn.close()
            except Exception as e: print(f"Could not connect to database. Error: {e}")
        elif choice == '3':
            try:
                conn = sqlite3.connect('payroll_database.db')
                date_str = input("\nEnter the Payroll Date for the Journal Entry (e.g., 07/15/2025): ")
                date_obj = datetime.strptime(date_str, '%m/%d/%Y')
                query_date = date_obj.strftime('%m/%d/%Y')
                create_journal_entry(conn, query_date)
                conn.close()
            except ValueError:
                print("Invalid date format. Please use MM/DD/YYYY.")
            except Exception as e:
                print(f"An error occurred: {e}")
        elif choice == '4': print("\nGoodbye!"); break
        else: print("\nInvalid choice. Please enter a number from 1 to 4.")