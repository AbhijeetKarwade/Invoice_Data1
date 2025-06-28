from flask import Flask, render_template, send_from_directory, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
import traceback
from datetime import datetime

app = Flask(__name__, static_folder='static', template_folder='static')

# Configuration
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if the file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_excel_file(file_stream):
    """Quick validation that the file is actually an Excel file"""
    try:
        # Just read the first few bytes to check the file signature
        header = file_stream.read(8)
        file_stream.seek(0)  # Reset file pointer
        
        # Check for Excel file signatures
        excel_signatures = [
            b'\x50\x4B\x05\x06',  # Empty ZIP (end of central directory)
            b'\x50\x4B\x03\x04',  # ZIP header
            b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'  # OLE2 (older Excel)
        ]
        
        return any(header.startswith(sig) for sig in excel_signatures)
    except Exception:
        return False

def process_excel_file(filepath):
    """Process Excel file with comprehensive error handling"""
    try:
        # First try with openpyxl (for .xlsx)
        try:
            df_custom = pd.read_excel(
                filepath,
                sheet_name='Custom Report',
                header=2,
                engine='openpyxl'
            )
            df_items = pd.read_excel(
                filepath,
                sheet_name='Item Details',
                header=2,
                engine='openpyxl'
            )
        except Exception as e:
            # Fall back to xlrd if openpyxl fails
            try:
                df_custom = pd.read_excel(
                    filepath,
                    sheet_name='Custom Report',
                    header=2,
                    engine='xlrd'
                )
                df_items = pd.read_excel(
                    filepath,
                    sheet_name='Item Details',
                    header=2,
                    engine='xlrd'
                )
            except Exception as fallback_error:
                raise ValueError(f"Failed to read Excel file with both openpyxl and xlrd: {str(fallback_error)}")

        # Basic data validation
        if df_custom.empty or df_items.empty:
            raise ValueError("One or both sheets are empty")

        # Clean column names
        df_custom.columns = df_custom.columns.str.strip()
        df_items.columns = df_items.columns.str.strip()

        # Check required columns
        required_columns = {
            'Custom Report': ['Date', 'Reference No'],
            'Item Details': ['Date', 'Invoice No./Txn No.']
        }
        
        for sheet, cols in required_columns.items():
            df = df_custom if sheet == 'Custom Report' else df_items
            missing = [col for col in cols if col not in df.columns]
            if missing:
                raise ValueError(f"Missing required columns in {sheet}: {', '.join(missing)}")

        # Process data (simplified version)
        processed_data = pd.concat([
            df_custom[['Date', 'Reference No']].rename(columns={'Date': 'date', 'Reference No': 'Ref_No'}),
            df_items[['Date', 'Invoice No./Txn No.']].rename(columns={'Date': 'date', 'Invoice No./Txn No.': 'Ref_No'})
        ])

        # Save processed file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        processed_filename = f"processed_{timestamp}_{secure_filename(os.path.basename(filepath))}"
        processed_path = os.path.join(app.config['PROCESSED_FOLDER'], processed_filename)
        
        processed_data.to_excel(processed_path, index=False, startrow=2)
        
        return processed_path, None

    except Exception as e:
        error_msg = f"Error processing file: {str(e)}\n{traceback.format_exc()}"
        app.logger.error(error_msg)
        return None, f"Processing error: {str(e)}"

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # Check if file exists in request
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in request'}), 400
        
        file = request.files['file']
        
        # Check if file was selected
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({'error': 'Only .xlsx files are allowed'}), 400
        
        # Validate it's actually an Excel file
        if not validate_excel_file(file.stream):
            return jsonify({'error': 'Invalid Excel file format'}), 400
        
        # Save the file temporarily
        filename = secure_filename(file.filename)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}")
        file.save(temp_path)
        
        # Process the file
        processed_path, error = process_excel_file(temp_path)
        
        # Clean up temp file
        try:
            os.remove(temp_path)
        except:
            pass
        
        if error:
            return jsonify({'error': error}), 500
        
        return jsonify({
            'message': 'File processed successfully',
            'processed_file': os.path.basename(processed_path)
        })

    except Exception as e:
        app.logger.error(f"Unexpected error in upload: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': 'Internal server error'}), 500

# ... (keep your other routes the same)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))