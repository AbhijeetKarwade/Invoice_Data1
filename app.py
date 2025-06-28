from flask import Flask, render_template, send_from_directory, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
import time
from threading import Thread
import logging

app = Flask(__name__, static_folder='static', template_folder='static')

# Configuration for file uploads
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB limit

# Ensure upload and processed directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_old_files():
    """Delete processed files older than 24 hours"""
    while True:
        try:
            now = time.time()
            for filename in os.listdir(app.config['PROCESSED_FOLDER']):
                filepath = os.path.join(app.config['PROCESSED_FOLDER'], filename)
                if os.path.isfile(filepath):
                    # Delete files older than 24 hours
                    if now - os.path.getmtime(filepath) > 86400:
                        os.remove(filepath)
                        app.logger.info(f"Deleted old file: {filename}")
        except Exception as e:
            app.logger.error(f"Error in cleanup: {str(e)}")
        time.sleep(3600)  # Run every hour

def process_excel_file(filepath):
    """Process the Excel file using the logic from app2.py"""
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

        # Handle missing values
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

        # Generate a unique filename for the processed file
        processed_filename = 'processed_' + secure_filename(os.path.basename(filepath))
        processed_path = os.path.join(app.config['PROCESSED_FOLDER'], processed_filename)

        # Save to a new Excel file with column names starting from the 3rd row
        parent_table.to_excel(processed_path, index=False, startrow=2)
        
        return processed_path, None
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
        
        try:
            # Process the Excel file
            processed_path, error = process_excel_file(filepath)
            
            # Delete the original uploaded file after processing
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    app.logger.info(f"Deleted uploaded file: {filename}")
            except Exception as e:
                app.logger.error(f"Error deleting uploaded file: {str(e)}")
            
            if error:
                return jsonify({'error': error}), 500
            
            return jsonify({
                'message': 'File successfully processed',
                'processed_file': os.path.basename(processed_path)
            })
        
        except Exception as e:
            # Ensure file is deleted even if processing fails
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    app.logger.info(f"Deleted uploaded file after processing failure: {filename}")
            except Exception as delete_error:
                app.logger.error(f"Error deleting uploaded file after processing failure: {str(delete_error)}")
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/processed/<filename>')
def processed_file(filename):
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename)

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File too large (max 50MB)'}), 413

# Start the cleanup thread when the app starts
if not app.debug or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
    Thread(target=cleanup_old_files, daemon=True).start()

if __name__ == '__main__':
    app.run(debug=True)