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

# Function to dynamically detect column names and search for the item
def process_excel(file_path, search_term):
    try:
        # Read the Excel file using pandas
        df = pd.read_excel(file_path)

        # Try to detect relevant columns dynamically
        expected_columns = ['Item', 'Description', 'Unit', 'Quantity', 'Unit Cost', 'Total Cost']
        actual_columns = df.columns

        # Find the best match for each expected column
        column_mapping = {}
        for expected_col in expected_columns:
            best_match = process.extractOne(expected_col, actual_columns, scorer=fuzz.token_sort_ratio)
            if best_match and best_match[1] > 70:  # Match confidence > 70%
                column_mapping[expected_col] = best_match[0]
        
        # Ensure we have the necessary columns for the search
        if 'Item' not in column_mapping and 'Description' not in column_mapping:
            return {'status': 'error', 'message': 'No recognizable item column found in the Excel file'}

        # Search for the term in the 'Item' or 'Description' column
        search_column = column_mapping.get('Item', column_mapping.get('Description'))
        filtered_rows = df[df[search_column].apply(lambda x: fuzz.partial_ratio(x, search_term) > 70)]

        if filtered_rows.empty:
            return {'status': 'success', 'message': 'No matching item found'}

        # Build the results to return
        results = []
        for _, row in filtered_rows.iterrows():
            item_info = {
                'Item': row.get(column_mapping.get('Item', column_mapping.get('Description')), 'N/A'),
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
