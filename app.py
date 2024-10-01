from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = './uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Route for uploading the file
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    # Process the uploaded file
    results = process_excel(file_path)
    
    return jsonify(results)

# Function to process the Excel file
def process_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        projects = []

        for _, row in df.iterrows():
            project_info = {
                'Project Name': row.get('Project Name', 'N/A'),
                'Square Footage': row.get('Square Footage', 'N/A'),
                'Unit Cost': row.get('Unit Cost', 'N/A')
            }
            projects.append(project_info)

        os.remove(file_path)
        return {'status': 'success', 'projects': projects}

    except Exception as e:
        return {'status': 'error', 'message': str(e)}

# Root route to display a simple welcome message or HTML form
@app.route('/')
def index():
    return '''
        <h1>Welcome to the Estimate Processor API</h1>
        <p>Use the form below to upload an Excel file:</p>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="Upload">
        </form>
    '''

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
