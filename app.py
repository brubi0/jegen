import pandas as pd
import re
from io import StringIO
from flask import Flask, request, render_template, make_response, session, redirect, url_for

# Initialize the Flask app
app = Flask(__name__)
app.secret_key = 'your_super_secret_key' # Replace with a real secret key

# --- Application Version ---
__version__ = "1.1.0"

# --- Constants ---
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

# --- Core Processing Functions ---
def process_payroll_register(df):
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

def create_journal_entry(register_df, taxes_df, payroll_date):
    if register_df.empty or taxes_df.empty:
        return None, "Error: Processed data is empty, cannot create Journal Entry.", 0, 0, False

    tax_map = {
        'EE_FICA': taxes_df[taxes_df['Description'].str.contains('Social Security - EE|Medicare - EE', regex=True)]['Numeric Value'].sum(),
        'ER_FICA': taxes_df[taxes_df['Description'].str.contains('Social Security - ER|Medicare - ER', regex=True)]['Numeric Value'].sum(),
        'FUTA': taxes_df[taxes_df['Description'] == 'Federal Unemployment Tax']['Numeric Value'].sum(),
        'SUTA': taxes_df[taxes_df['Description'] == 'State Unemployment/Disability Ins - ER']['Numeric Value'].sum(),
        'State_WH': taxes_df[taxes_df['Description'] == 'State Income Tax']['Numeric Value'].sum(),
        'SDI': taxes_df[taxes_df['Description'] == 'State Disability Insurance - EE']['Numeric Value'].sum()
    }
    
    journal_lines = []
    entity = 2
    formatted_date = pd.to_datetime(payroll_date).strftime('%m/%d/%Y')

    for index, employee in register_df.iterrows():
        emp_name, emp_dept = employee['Employee Name'], employee['Department']
        if employee['Gross Pay'] > 0:
            acct = CHART_OF_ACCOUNTS['Gross Pay']
            journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Gross Pay: {emp_name}", 'Department': emp_dept, 'Debit': employee['Gross Pay'], 'Credit': 0, 'Subsidiary': entity})
        if employee['Commission'] > 0:
            acct = CHART_OF_ACCOUNTS['Commission']
            journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Commission: {emp_name}", 'Department': emp_dept, 'Debit': employee['Commission'], 'Credit': 0, 'Subsidiary': entity})
        if employee['Car Allowance'] > 0:
            acct = CHART_OF_ACCOUNTS['Car Allowance']
            journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Car Allowance: {emp_name}", 'Department': emp_dept, 'Debit': employee['Car Allowance'], 'Credit': 0, 'Subsidiary': entity})

    acct = CHART_OF_ACCOUNTS['ER_FICA']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer FICA', 'Department': '', 'Debit': tax_map['ER_FICA'], 'Credit': 0, 'Subsidiary': entity})
    acct = CHART_OF_ACCOUNTS['FUTA']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer FUTA', 'Department': '', 'Debit': tax_map['FUTA'], 'Credit': 0, 'Subsidiary': entity})
    acct = CHART_OF_ACCOUNTS['SUTA']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employer SUTA', 'Department': '', 'Debit': tax_map['SUTA'], 'Credit': 0, 'Subsidiary': entity})
    acct = CHART_OF_ACCOUNTS['EE_FICA']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employee FICA Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['EE_FICA'], 'Subsidiary': entity})
    acct = CHART_OF_ACCOUNTS['State_WH']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'State Income Tax Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['State_WH'], 'Subsidiary': entity})
    acct = CHART_OF_ACCOUNTS['SDI']; journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': 'Employee SDI Withheld', 'Department': '', 'Debit': 0, 'Credit': tax_map['SDI'], 'Subsidiary': entity})

    je_df_temp = pd.DataFrame(journal_lines)
    net_pay = je_df_temp['Debit'].sum() - je_df_temp['Credit'].sum()
    acct = CHART_OF_ACCOUNTS['Net_Pay']
    journal_lines.append({'Date': formatted_date, 'Account': f"{acct['acct']} {acct['desc']}", 'Memo': f"Payroll Cash Clearing for {formatted_date}", 'Department': '', 'Debit': 0, 'Credit': net_pay, 'Subsidiary': entity})

    final_je_df = pd.DataFrame(journal_lines)
    final_je_df[['Debit', 'Credit']] = final_je_df[['Debit', 'Credit']].round(2)
    
    column_order = ['Date', 'Account', 'Memo', 'Department', 'Debit', 'Credit', 'Subsidiary']
    final_je_df = final_je_df[column_order]

    total_debits = final_je_df['Debit'].sum()
    total_credits = final_je_df['Credit'].sum()
    is_balanced = abs(total_debits - total_credits) < 0.01

    if not is_balanced:
        status_message = f"WARNING: Journal Entry is out of balance! Debits: ${total_debits:,.2f}, Credits: ${total_credits:,.2f}"
    else:
        status_message = f"Success! Debits and Credits are balanced: ${total_debits:,.2f}"
    
    return final_je_df, status_message, total_debits, total_credits, is_balanced

# --- FLASK WEB ROUTES ---
@app.route('/')
def index():
    return render_template('index.html', version=__version__)

@app.route('/process', methods=['POST'])
def process_files_route():
    payroll_date = request.form.get('payroll_date')
    payroll_register_file = request.files.get('payroll_register')
    statistical_summary_file = request.files.get('statistical_summary')

    if not all([payroll_date, payroll_register_file, statistical_summary_file]):
        return "Error: Please provide all inputs.", 400
    
    try:
        df_register_raw = pd.read_excel(payroll_register_file, sheet_name='Payroll Register', header=None)
        df_summary_raw = pd.read_excel(statistical_summary_file, sheet_name='Statistical Summary', header=None)

        processed_register_df = process_payroll_register(df_register_raw)
        processed_taxes_df = process_statistical_summary(df_summary_raw)

        final_je_df, status_message, total_debits, total_credits, is_balanced = create_journal_entry(processed_register_df, processed_taxes_df, payroll_date)

        if final_je_df is None:
            return f"An error occurred: {status_message}", 500

        # Store all necessary data in the session
        session['journal_entry_data'] = final_je_df.to_json(orient='split')
        session['status_message'] = status_message
        session['total_debits'] = total_debits
        session['total_credits'] = total_credits
        # Convert boolean to int (1 or 0) for session serialization
        session['is_balanced'] = 1 if is_balanced else 0
        
        return redirect(url_for('review_page'))

    except Exception as e:
        return f"An error occurred during processing: {e}", 500

@app.route('/review')
def review_page():
    je_data_json = session.get('journal_entry_data')
    if not je_data_json:
        return redirect(url_for('index'))
    
    df = pd.read_json(StringIO(je_data_json), orient='split')
    data_for_template = df.to_dict(orient='records')

    # Pass all the session data to the template
    return render_template('review.html', 
                           je_data=data_for_template, 
                           version=__version__,
                           status_message=session.get('status_message'),
                           total_debits=session.get('total_debits'),
                           total_credits=session.get('total_credits'),
                           # Convert int (1 or 0) back to boolean
                           is_balanced=bool(session.get('is_balanced')))

@app.route('/generate-csv', methods=['POST'])
def generate_csv():
    """Takes edited data as a JSON string and generates the final CSV."""
    edited_data_json = request.form.get('edited_data')
    if not edited_data_json:
        return "Error: No data received.", 400

    edited_df = pd.read_json(StringIO(edited_data_json))

    column_order = ['Date', 'Account', 'Memo', 'Department', 'Debit', 'Credit', 'Subsidiary']
    edited_df = edited_df.reindex(columns=column_order)

    csv_df = edited_df.copy()
    if 'Debit' in csv_df.columns:
        csv_df['Debit'] = pd.to_numeric(csv_df['Debit'], errors='coerce').fillna(0)
        csv_df['Debit'] = csv_df['Debit'].apply(lambda x: '' if x == 0 else x)
    if 'Credit' in csv_df.columns:
        csv_df['Credit'] = pd.to_numeric(csv_df['Credit'], errors='coerce').fillna(0)
        csv_df['Credit'] = csv_df['Credit'].apply(lambda x: '' if x == 0 else x)
    
    response = make_response(csv_df.to_csv(index=False))
    response.headers["Content-Disposition"] = "attachment; filename=journal_entry_edited.csv"
    response.headers["Content-Type"] = "text/csv"
    
    return response

if __name__ == '__main__':
    app.run(debug=True)