from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = './uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    results = process_excel(file_path)

    return jsonify(results)

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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
