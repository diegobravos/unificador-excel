import io
import base64
import pandas as pd
from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__, template_folder='templates')


def read_file_from_bytes(file_bytes, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    bio = io.BytesIO(file_bytes)
    if ext == 'csv':
        return pd.read_csv(bio, dtype=str)
    return pd.read_excel(bio, dtype=str)


def extract_header_styles(file_bytes, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'csv':
        return {}
    styles = {}
    try:
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        for cell in ws[1]:
            col_name = str(cell.value) if cell.value is not None else ''
            if not col_name:
                continue
            style = {}
            try:
                if cell.fill and cell.fill.fill_type not in (None, 'none'):
                    fg = cell.fill.fgColor
                    if fg.type == 'rgb' and fg.rgb not in ('00000000', 'FFFFFFFF', '00FFFFFF'):
                        style['bg_color'] = fg.rgb
            except Exception:
                pass
            try:
                if cell.font:
                    style['bold'] = bool(cell.font.bold)
                    if cell.font.color and cell.font.color.type == 'rgb':
                        fc = cell.font.color.rgb
                        if fc not in ('FF000000', '00000000'):
                            style['font_color'] = fc
            except Exception:
                pass
            try:
                col_letter = cell.column_letter
                dim = ws.column_dimensions.get(col_letter)
                if dim and dim.width:
                    style['width'] = dim.width
            except Exception:
                pass
            styles[col_name] = style
        wb.close()
    except Exception:
        pass
    return styles


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    files = request.files.getlist('files')
    all_columns = set()
    all_styles = {}
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
            # Collect styles; first file wins for each column
            styles = extract_header_styles(file_bytes, f.filename)
            for col, style in styles.items():
                if col not in all_styles:
                    all_styles[col] = style
        except Exception as e:
            return jsonify({'error': f'Error leyendo {f.filename}: {str(e)}'}), 400

    if not file_data_list:
        return jsonify({'error': 'Ningún archivo válido fue cargado'}), 400

    return jsonify({
        'columns': sorted(list(all_columns)),
        'column_styles': all_styles,
        'files': file_info,
        'file_data': file_data_list
    })


@app.route('/merge', methods=['POST'])
def merge_files():
    data = request.get_json()
    column_order   = data.get('columns', [])   # ordered list
    column_styles  = data.get('column_styles', {})
    dedup_column   = data.get('dedup_column') or None
    file_data_list = data.get('file_data', [])

    if not column_order or not file_data_list:
        return jsonify({'error': 'Faltan parámetros'}), 400

    dfs = []
    for fd in file_data_list:
        try:
            file_bytes = base64.b64decode(fd['data'])
            df = read_file_from_bytes(file_bytes, fd['name'])
            cols_to_use = [c for c in column_order if c in df.columns]
            dfs.append(df[cols_to_use])
        except Exception as e:
            return jsonify({'error': f'Error procesando {fd["name"]}: {str(e)}'}), 400

    merged = pd.concat(dfs, ignore_index=True)

    # Reorder columns to match requested order
    final_cols = [c for c in column_order if c in merged.columns]
    merged = merged[final_cols]

    before = len(merged)
    if dedup_column and dedup_column in merged.columns:
        merged = merged.drop_duplicates(subset=[dedup_column])
    else:
        merged = merged.drop_duplicates()
    duplicates_removed = before - len(merged)

    # Build output with openpyxl to preserve styles
    wb = Workbook()
    ws = wb.active

    # Write header row with original styles
    for col_idx, col_name in enumerate(final_cols, 1):
        cell  = ws.cell(row=1, column=col_idx, value=col_name)
        style = column_styles.get(col_name, {})

        if 'bg_color' in style:
            cell.fill = PatternFill('solid', fgColor=style['bg_color'])

        font_kwargs = {'bold': style.get('bold', True)}
        if 'font_color' in style:
            font_kwargs['color'] = style['font_color']
        cell.font = Font(**font_kwargs)
        cell.alignment = Alignment(horizontal='left', vertical='center')

    # Write data rows
    for row_idx, row in enumerate(merged.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx,
                    value='' if (value is None or (isinstance(value, float) and pd.isna(value))) else value)

    # Set column widths
    for col_idx, col_name in enumerate(final_cols, 1):
        col_letter = get_column_letter(col_idx)
        style = column_styles.get(col_name, {})
        ws.column_dimensions[col_letter].width = style.get('width', max(len(col_name) + 4, 14))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    output_b64 = base64.b64encode(output.read()).decode('utf-8')

    return jsonify({
        'output_data': output_b64,
        'total_rows': len(merged),
        'duplicates_removed': duplicates_removed,
        'columns': final_cols
    })


if __name__ == '__main__':
    import os
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
