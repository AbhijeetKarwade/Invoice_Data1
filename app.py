from flask import Flask, render_template, send_from_directory, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='static', template_folder='static')

# Configuration for file uploads
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload
app.config['UPLOAD_TIMEOUT'] = 300

# Ensure upload and processed directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel_file(filepath):
    """Optimized Excel processing with memory efficiency"""
    try:
        # Read sheets with optimized parameters
        chunksize = 5000  # Process in chunks for memory efficiency
        reader = pd.read_excel(filepath, sheet_name=None, header=2, chunksize=chunksize)
        
        # Process each sheet
        dfs = {}
        for sheet_name, chunk in reader.items():
            dfs[sheet_name] = chunk
            
        # Get the required sheets
        df_custom = dfs.get('Custom Report', pd.DataFrame())
        df_items = dfs.get('Item Details', pd.DataFrame())

        # Clean column names
        df_custom.columns = df_custom.columns.str.strip()
        df_items.columns = df_items.columns.str.strip()

        # Select key columns and rename
        df_custom_key = df_custom[['Date', 'Reference No']].rename(columns={
            'Date': 'date', 
            'Reference No': 'Ref_No'
        })
        df_items_key = df_items[['Date', 'Invoice No./Txn No.']].rename(columns={
            'Date': 'date', 
            'Invoice No./Txn No.': 'Ref_No'
        })

        # Combine key columns
        parent_table = pd.concat([df_custom_key, df_items_key], ignore_index=True)
        parent_table['date'] = pd.to_datetime(parent_table['date'], dayfirst=True, errors='coerce')

        # Process remaining columns
        custom_remaining = df_custom.drop(columns=['Date', 'Reference No'], errors='ignore')
        items_remaining = df_items.drop(columns=['Date', 'Invoice No./Txn No.'], errors='ignore')

        # Add prefixes
        custom_remaining.columns = [f'Custom_{col}' for col in custom_remaining.columns]
        items_remaining.columns = [f'Items_{col}' for col in items_remaining.columns]

        # Add back reference columns
        custom_remaining['Ref_No'] = df_custom['Reference No']
        custom_remaining['date'] = df_custom['Date']
        items_remaining['Ref_No'] = df_items['Invoice No./Txn No.']
        items_remaining['date'] = df_items['Date']

        # Convert dates
        custom_remaining['date'] = pd.to_datetime(custom_remaining['date'], dayfirst=True, errors='coerce')
        items_remaining['date'] = pd.to_datetime(items_remaining['date'], dayfirst=True, errors='coerce')

        # Merge data
        parent_table = parent_table.merge(custom_remaining, on=['Ref_No', 'date'], how='outer')
        parent_table = parent_table.merge(items_remaining, on=['Ref_No', 'date'], how='outer')

        # Handle missing values
        string_cols = [col for col in parent_table.columns if parent_table[col].dtype == 'object']
        numeric_cols = [col for col in parent_table.columns if parent_table[col].dtype in ['float64', 'int64']]
        parent_table[string_cols] = parent_table[string_cols].fillna('')
        parent_table[numeric_cols] = parent_table[numeric_cols].fillna(0)

        # Sort and format
        parent_table = parent_table.sort_values(by=['date', 'Ref_No'])
        parent_table['date'] = parent_table['date'].dt.strftime('%d/%m/%Y')
        parent_table = parent_table.drop_duplicates(keep='first')

        # Save processed file
        processed_filename = 'processed_' + secure_filename(os.path.basename(filepath))
        processed_path = os.path.join(app.config['PROCESSED_FOLDER'], processed_filename)
        
        # Optimized saving
        with pd.ExcelWriter(processed_path, engine='openpyxl') as writer:
            parent_table.to_excel(writer, index=False, startrow=2)
        
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
        
        processed_path, error = process_excel_file(filepath)
        
        if error:
            # Clean up the uploaded file if processing fails
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'error': error}), 500
        
        return jsonify({
            'message': 'File successfully processed',
            'processed_file': os.path.basename(processed_path)
        })
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/processed/<filename>')
def processed_file(filename):
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename)

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

if __name__ == '__main__':
    app.run(debug=True)