from flask import Flask, request, redirect, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)

# Set the folder for file uploads and results
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

# Ensure the directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'contacts' not in request.files or 'balances' not in request.files:
        return "Missing files", 400

    contacts_file = request.files['contacts']
    balances_file = request.files['balances']

    # Save the uploaded files to the server
    contacts_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'contacts.xlsx')
    balances_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'balances.xlsx')
    contacts_file.save(contacts_file_path)
    balances_file.save(balances_file_path)

    # Load the Excel files
    contacts_data = pd.read_excel(contacts_file_path)
    balances_data = pd.read_excel(balances_file_path)

    # Standardize column names for merging
    contacts_data.rename(columns={'AdmNo': 'Adm No.'}, inplace=True)

    # Merge the two dataframes on 'Adm No.'
    merged_data = pd.merge(balances_data, contacts_data[['Adm No.', 'Parent Contact']], on='Adm No.', how='left')

    # Save the result to a new Excel file
    result_file_path = os.path.join(app.config['RESULT_FOLDER'], 'balances_with_contacts.xlsx')
    merged_data.to_excel(result_file_path, index=False)

    return send_file(result_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
