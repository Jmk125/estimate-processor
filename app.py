from flask import Flask, request, jsonify
import pandas as pd
import os

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

# Function to process the Excel file and search for the item
def process_excel(file_path, search_term):
    try:
        # Read the Excel file using pandas
        df = pd.read_excel(file_path)

        # Assuming columns like 'Item', 'Unit', 'Quantity', 'Unit Cost', 'Total Cost'
        # You may need to adjust the column names to match your Excel file structure
        filtered_rows = df[df['Item'].str.contains(search_term, case=False, na=False)]

        if filtered_rows.empty:
            return {'status': 'success', 'message': 'No matching item found'}

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
