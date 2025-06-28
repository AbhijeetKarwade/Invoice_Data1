from flask import Flask, render_template, send_from_directory, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
import json
import math

app = Flask(__name__, static_folder='static', template_folder='static')

# Configuration for file uploads
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_nans(obj):
    """Convert NaN values to None (which becomes null in JSON)"""
    if isinstance(obj, float) and math.isnan(obj):
        return None
    elif isinstance(obj, dict):
        return {k: convert_nans(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_nans(item) for item in obj]
    return obj

def process_excel_file(filepath):
    """Process the Excel file and return the processed data in memory"""
    try:
        # Read the two sheets, specifying header row (row 3, so header=2)
        df_custom = pd.read_excel(filepath, sheet_name='Custom Report', header=2)
        df_items = pd.read_excel(filepath, sheet_name='Item Details', header=2)

        # Clean column names by stripping whitespace
        df_custom.columns = df_custom.columns.str.strip()
        df_items.columns = df_items.columns.str.strip()

        # Select key columns and rename for consistency
        df_custom_key = df_custom[['Date', 'Reference No']].rename(columns={'Date': 'date', 'Reference No': 'Ref_No'})
        df_items_key = df_items[['Date', 'Invoice No./Txn No.']].rename(columns={'Date': 'date', 'Invoice No./Txn No.': 'Ref_No'})

        # Combine key columns without removing duplicates
        parent_table = pd.concat([df_custom_key, df_items_key], ignore_index=True)
        parent_table['date'] = pd.to_datetime(parent_table['date'], dayfirst=True, errors='coerce')

        # Prepare remaining columns from both sheets
        custom_remaining = df_custom.drop(columns=['Date', 'Reference No'])
        items_remaining = df_items.drop(columns=['Date', 'Invoice No./Txn No.'])

        # Add prefixes to distinguish columns
        custom_remaining.columns = [f'Custom_{col}' for col in custom_remaining.columns]
        items_remaining.columns = [f'Items_{col}' for col in items_remaining.columns]

        # Add Ref_No and date back for merging
        custom_remaining['Ref_No'] = df_custom['Reference No']
        custom_remaining['date'] = df_custom['Date']
        items_remaining['Ref_No'] = df_items['Invoice No./Txn No.']
        items_remaining['date'] = df_items['Date']

        # Convert dates to datetime for consistent merging
        custom_remaining['date'] = pd.to_datetime(custom_remaining['date'], dayfirst=True, errors='coerce')
        items_remaining['date'] = pd.to_datetime(items_remaining['date'], dayfirst=True, errors='coerce')

        # Merge with parent_table using Ref_No and date to preserve duplicates
        parent_table = parent_table.merge(custom_remaining, on=['Ref_No', 'date'], how='outer')
        parent_table = parent_table.merge(items_remaining, on=['Ref_No', 'date'], how='outer')

        # Handle missing values - fill with empty string for strings, 0 for numbers
        string_columns = [col for col in parent_table.columns if parent_table[col].dtype == 'object']
        numeric_columns = [col for col in parent_table.columns if parent_table[col].dtype in ['float64', 'int64']]
        parent_table[string_columns] = parent_table[string_columns].fillna('')
        parent_table[numeric_columns] = parent_table[numeric_columns].fillna(0)

        # Sort by date and Ref_No
        parent_table = parent_table.sort_values(by=['date', 'Ref_No'])

        # Format date to DD/MM/YYYY
        parent_table['date'] = parent_table['date'].dt.strftime('%d/%m/%Y')

        # Remove duplicates based on all columns
        parent_table = parent_table.drop_duplicates(keep='first')

        # Convert the processed DataFrame to a dictionary and clean NaN values
        processed_data = parent_table.to_dict(orient='records')
        processed_data = convert_nans(processed_data)
        
        return processed_data, list(parent_table.columns)
    except Exception as e:
        return None, str(e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/print')
def print_view():
    return render_template('print.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Process the Excel file and get data in memory
        processed_data, error_or_headers = process_excel_file(filepath)
        
        # Remove the uploaded file after processing to save space
        os.remove(filepath)
        
        if isinstance(error_or_headers, str):  # It's an error message
            return jsonify({'error': error_or_headers}), 500
        
        return jsonify({
            'message': 'File successfully processed',
            'data': processed_data,
            'headers': error_or_headers  # In this case, it's the headers list
        })
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

if __name__ == '__main__':
    app.run(debug=True)