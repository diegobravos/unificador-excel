import io
import base64
import pandas as pd
from flask import Flask, request, jsonify, render_template

app = Flask(__name__, template_folder='templates')


def read_file_from_bytes(file_bytes, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    bio = io.BytesIO(file_bytes)
    if ext == 'csv':
        return pd.read_csv(bio, dtype=str)
    return pd.read_excel(bio, dtype=str)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    files = request.files.getlist('files')
    all_columns = set()
    file_data_list = []
    file_info = []

    for f in files:
        if not f or not f.filename:
            continue
        ext = f.filename.rsplit('.', 1)[1].lower()
        if ext not in ('xlsx', 'xls', 'csv'):
            continue

        file_bytes = f.read()
        try:
            df = read_file_from_bytes(file_bytes, f.filename)
            cols = list(df.columns)
            all_columns.update(cols)
            file_info.append({'name': f.filename, 'rows': len(df), 'columns': cols})
            file_data_list.append({
                'name': f.filename,
                'data': base64.b64encode(file_bytes).decode('utf-8')
            })
        except Exception as e:
            return jsonify({'error': f'Error leyendo {f.filename}: {str(e)}'}), 400

    if not file_data_list:
        return jsonify({'error': 'Ningún archivo válido fue cargado'}), 400

    return jsonify({
        'columns': sorted(list(all_columns)),
        'files': file_info,
        'file_data': file_data_list
    })


@app.route('/merge', methods=['POST'])
def merge_files():
    data = request.get_json()
    selected_columns = data.get('columns', [])
    dedup_column = data.get('dedup_column') or None
    file_data_list = data.get('file_data', [])

    if not selected_columns or not file_data_list:
        return jsonify({'error': 'Faltan parámetros'}), 400

    dfs = []
    for fd in file_data_list:
        try:
            file_bytes = base64.b64decode(fd['data'])
            df = read_file_from_bytes(file_bytes, fd['name'])
            cols_to_use = [c for c in selected_columns if c in df.columns]
            dfs.append(df[cols_to_use])
        except Exception as e:
            return jsonify({'error': f'Error procesando {fd["name"]}: {str(e)}'}), 400

    merged = pd.concat(dfs, ignore_index=True)
    before = len(merged)

    if dedup_column and dedup_column in merged.columns:
        merged = merged.drop_duplicates(subset=[dedup_column])
    else:
        merged = merged.drop_duplicates()

    duplicates_removed = before - len(merged)

    output = io.BytesIO()
    merged.to_excel(output, index=False)
    output.seek(0)
    output_b64 = base64.b64encode(output.read()).decode('utf-8')

    return jsonify({
        'output_data': output_b64,
        'total_rows': len(merged),
        'duplicates_removed': duplicates_removed,
        'columns': list(merged.columns)
    })


if __name__ == '__main__':
    import os
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
