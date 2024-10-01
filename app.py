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

        # Try to load the detail sheet, skipping initial rows to locate the column headers
        for skiprows in range(0, 10):  # Try skipping between 0 to 10 rows
            df = pd.read_excel(file_path, sheet_name=detail_sheet_name, skiprows=skiprows)
            if 'Unnamed' not in df.columns[0]:  # Check if we have usable columns
                break

        # Fallback to positional columns if necessary
        if 'Unnamed' in df.columns[0] or df.empty:
            # If no valid column names are found, fallback to column index-based processing
            df = pd.read_excel(file_path, sheet_name=detail_sheet_name, skiprows=10, header=None)
            # Assuming Column C = Items, D = Quantity, E = Unit, F = Unit Cost, G = Total Cost
            df = df.iloc[:, [2, 3, 4, 5, 6]]  # Only keep columns C, D, E, F, G
            df.columns = ['Item', 'Quantity', 'Unit', 'Unit Cost', 'Total Cost']

        # Now, perform a fuzzy search on the "Item" column (or assumed column C)
        filtered_rows = df[df['Item'].apply(lambda x: fuzz.partial_ratio(str(x), search_term) > 70)]

        if filtered_rows.empty:
            return {'status': 'success', 'message': 'No matching item found'}

        # Build the results to return
        results = []
        for _, row in filtered_rows.iterrows():
            item_info = {
                'Item': row.get('Item', 'N/A'),
                'Unit': row.get('Unit', 'N/A'),
                'Quantity': row.get('Quantity', 'N/A'),
                'Unit Cost': row.get('Unit Cost', 'N/A'),
                'Total Cost': row.get('Total Cost', 'N/A')
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
