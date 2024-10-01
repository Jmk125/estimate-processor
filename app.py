from flask import Flask, request, jsonify
import pandas as pd
import os
from fuzzywuzzy import fuzz, process  # For fuzzy string matching

app = Flask(__name__)

UPLOAD_FOLDER = './uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Route for uploading the file and searching
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files or 'search_term' not in request.form:
        return jsonify({'error': 'File or search term missing'}), 400

    file = request.files['file']
    search_term = request.form['search_term']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    # Process the uploaded file and search for the term
    results = process_excel(file_path, search_term)

    return jsonify(results)

# Function to dynamically detect the correct sheet and columns, then search for the item
def process_excel(file_path, search_term):
    try:
        # Read the Excel file with all sheets
        xls = pd.ExcelFile(file_path)

        # Find the sheet that contains 'Detail' in its name
        detail_sheet_name = None
        for sheet_name in xls.sheet_names:
            if 'Detail' in sheet_name:
                detail_sheet_name = sheet_name
                break
        
        if not detail_sheet_name:
            return {'status': 'error', 'message': 'No sheet containing "Detail" found'}

        # Load the detail sheet into a dataframe
        df = pd.read_excel(file_path, sheet_name=detail_sheet_name)

        # Debug: Print available columns to help identify the issue
        print("Available columns:", df.columns.tolist())

        # Try to detect relevant columns dynamically
        # Expected columns based on user's description
        expected_columns = {
            'Item': ['Item', 'Description'],  # Items are typically in column C, with names like "Item" or "Description"
            'Quantity': ['Quantity'],         # Quantities in column D
            'Unit': ['Unit'],                 # Units in column E
            'Unit Cost': ['Unit Cost'],       # Unit cost in column F
            'Total Cost': ['Total Cost']      # Total cost in column G
        }

        # Map the actual columns from the file to the expected columns
        column_mapping = {}
        actual_columns = df.columns

        for key, expected_col_names in expected_columns.items():
            best_match = process.extractOne(expected_col_names[0], actual_columns, scorer=fuzz.token_sort_ratio)
            if best_match and best_match[1] > 70:  # Match confidence > 70%
                column_mapping[key] = best_match[0]

        # Debug: Print the column mapping for review
        print("Column mapping:", column_mapping)

        # Check if we found the required columns
        if 'Item' not in column_mapping:
            return {'status': 'error', 'message': 'No recognizable item column found in the Excel file'}

        # Search for the term in the 'Item' column
        search_column = column_mapping['Item']
        filtered_rows = df[df[search_column].apply(lambda x: fuzz.partial_ratio(str(x), search_term) > 70)]

        if filtered_rows.empty:
            return {'status': 'success', 'message': 'No matching item found'}

        # Build the results to return
        results = []
        for _, row in filtered_rows.iterrows():
            item_info = {
                'Item': row.get(column_mapping.get('Item', 'N/A'), 'N/A'),
                'Unit': row.get(column_mapping.get('Unit', 'N/A')),
                'Quantity': row.get(column_mapping.get('Quantity', 'N/A')),
                'Unit Cost': row.get(column_mapping.get('Unit Cost', 'N/A')),
                'Total Cost': row.get(column_mapping.get('Total Cost', 'N/A'))
            }
            results.append(item_info)

        # Delete the file after processing
        os.remove(file_path)

        return {'status': 'success', 'items': results}

    except Exception as e:
        return {'status': 'error', 'message': str(e)}

# Root route to display the HTML form for file upload and searching
@app.route('/')
def index():
    return '''
        <h1>Upload Excel File and Search for an Item</h1>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <label for="file">Choose Excel file:</label>
            <input type="file" name="file" required><br><br>

            <label for="search_term">Enter search term (e.g., Item Name):</label>
            <input type="text" name="search_term" required><br><br>

            <input type="submit" value="Upload and Search">
        </form>
    '''

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
