from flask import Flask, request, jsonify
import pandas as pd
import os
from fuzzywuzzy import fuzz

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

# Function to search the entire spreadsheet for fuzzy matches to the search term
def process_excel(file_path, search_term):
    try:
        # Read the Excel file with all sheets
        xls = pd.ExcelFile(file_path)

        # Iterate over each sheet in the Excel file
        all_matches = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Convert all cells to strings for consistent processing
            df = df.astype(str)

            # Iterate over all rows and cells to find fuzzy matches
            for idx, row in df.iterrows():
                for col_name, cell_value in row.items():
                    if fuzz.partial_ratio(str(cell_value), search_term) > 70:
                        # Store the entire row if we find a fuzzy match
                        match_info = {
                            'Sheet': sheet_name,
                            'Row': idx + 1,  # Excel rows are 1-indexed
                            'Matched Cell': col_name,
                            'Matched Value': cell_value,
                            'Row Data': row.to_dict()  # Return entire row for context
                        }
                        all_matches.append(match_info)

        if not all_matches:
            return {'status': 'success', 'message': 'No matching item found'}

        return {'status': 'success', 'matches': all_matches}

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
