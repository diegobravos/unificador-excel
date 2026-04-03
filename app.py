import io
import re
import base64
import unicodedata
import pandas as pd
from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

import json
import math

app = Flask(__name__, template_folder='templates')


class _SafeEncoder(json.JSONEncoder):
    """Converts NaN/Inf floats to None so the response is always valid JSON."""
    def default(self, obj):
        return super().default(obj)

    def iterencode(self, o, _one_shot=False):
        # Replace NaN and Infinity before encoding
        return super().iterencode(self._clean(o), _one_shot)

    def _clean(self, obj):
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
        elif isinstance(obj, dict):
            return {k: self._clean(v) for k, v in obj.items()}
        elif isinstance(obj, (list, tuple)):
            return [self._clean(v) for v in obj]
        return obj


app.json.encoder = _SafeEncoder


# ── File reading ───────────────────────────────────────────────────────────────

def read_file_from_bytes(file_bytes, filename, sheet_name=None):
    ext = filename.rsplit('.', 1)[1].lower()
    bio = io.BytesIO(file_bytes)
    if ext == 'csv':
        return pd.read_csv(bio, dtype=str)
    kwargs = {'dtype': str}
    if sheet_name:
        kwargs['sheet_name'] = sheet_name
    return pd.read_excel(bio, **kwargs)


def read_all_sheets_info(file_bytes, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'csv':
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
        return {'(hoja única)': {'columns': list(df.columns), 'rows': len(df)}}
    try:
        xl  = pd.ExcelFile(io.BytesIO(file_bytes))
        out = {}
        for sn in xl.sheet_names:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sn, dtype=str)
            out[sn] = {'columns': list(df.columns), 'rows': len(df)}
        return out
    except Exception:
        return {}


def extract_header_styles(file_bytes, filename, sheet_name=None):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'csv':
        return {}
    styles = {}
    try:
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
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


# ── Analysis helpers ───────────────────────────────────────────────────────────

def normalize_string(s):
    """Lowercase, remove accents, remove punctuation, collapse spaces."""
    s = str(s).lower().strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def find_similar_values(df, columns):
    """Return per-column groups of values that normalize to the same string."""
    suggestions = []
    for col in columns:
        if col not in df.columns:
            continue
        series = df[col].dropna().astype(str)
        if series.empty:
            continue
        # Skip purely numeric columns
        if pd.to_numeric(series, errors='coerce').notna().all():
            continue
        unique_vals = series.unique()
        if len(unique_vals) < 2 or len(unique_vals) > 300:
            continue

        # Group by normalized form
        norm_map = {}
        for v in unique_vals:
            key = normalize_string(v)
            norm_map.setdefault(key, []).append(v)

        col_groups = []
        for variants in norm_map.values():
            unique_variants = list(dict.fromkeys(variants))  # dedupe, preserve order
            if len(unique_variants) < 2:
                continue
            # Canonical = most frequent value
            counts = series.value_counts()
            canonical = max(unique_variants, key=lambda v: counts.get(v, 0))
            col_groups.append({'canonical': canonical, 'variants': unique_variants})

        if col_groups:
            suggestions.append({'column': col, 'groups': col_groups})
    return suggestions


def find_duplicate_groups(df, dedup_column, max_groups=50):
    """Return list of duplicate groups with their row ids and source files."""
    data_cols = [c for c in df.columns if c not in ('__source__', '__row_id__')]
    groups = []

    if dedup_column and dedup_column in df.columns:
        dup_mask = df.duplicated(subset=[dedup_column], keep=False)
        if not dup_mask.any():
            return []
        for key_val in df.loc[dup_mask, dedup_column].unique()[:max_groups]:
            mask    = df[dedup_column] == key_val
            grp_df  = df[mask]
            groups.append({
                'key_col': dedup_column,
                'key_val': str(key_val),
                'indices': grp_df['__row_id__'].tolist(),
                'records': grp_df[data_cols].fillna('').to_dict('records'),
                'sources': grp_df['__source__'].tolist() if '__source__' in df.columns else [],
            })
    else:
        # Full-row duplicates
        df_hash = df[data_cols].copy()
        df_hash['__hash__'] = df_hash.apply(
            lambda r: hash(tuple(str(x) for x in r.values)), axis=1)
        dup_mask = df_hash.duplicated(subset=['__hash__'], keep=False)
        if not dup_mask.any():
            return []
        seen = set()
        for _, row in df_hash[dup_mask].iterrows():
            h = row['__hash__']
            if h in seen or len(groups) >= max_groups:
                continue
            seen.add(h)
            same = df[df_hash['__hash__'] == h]
            groups.append({
                'key_col': None,
                'key_val': f'dup_{len(groups)}',
                'indices': same['__row_id__'].tolist(),
                'records': same[data_cols].fillna('').to_dict('records'),
                'sources': same['__source__'].tolist() if '__source__' in df.columns else [],
            })
    return groups


def _build_merged(file_data_list, column_order, sheet_selection, priority_file):
    """Common merge logic used by /preview and /confirm."""
    if priority_file:
        pri    = [fd for fd in file_data_list if fd['name'] == priority_file]
        others = [fd for fd in file_data_list if fd['name'] != priority_file]
        file_data_list = pri + others

    dfs = []
    for fd in file_data_list:
        file_bytes  = base64.b64decode(fd['data'])
        sheet_name  = sheet_selection.get(fd['name']) or None
        df          = read_file_from_bytes(file_bytes, fd['name'], sheet_name=sheet_name)
        cols_to_use = [c for c in column_order if c in df.columns]
        sub         = df[cols_to_use].copy()
        sub['__source__'] = fd['name']
        dfs.append(sub)

    merged = pd.concat(dfs, ignore_index=True)
    merged['__row_id__'] = range(len(merged))
    return merged


def _write_excel(merged_df, final_cols, column_styles):
    wb = Workbook()
    ws = wb.active

    for col_idx, col_name in enumerate(final_cols, 1):
        cell  = ws.cell(row=1, column=col_idx, value=col_name)
        style = column_styles.get(col_name, {})
        if 'bg_color' in style:
            cell.fill = PatternFill('solid', fgColor=style['bg_color'])
        font_kwargs = {'bold': style.get('bold', True)}
        if 'font_color' in style:
            font_kwargs['color'] = style['font_color']
        cell.font      = Font(**font_kwargs)
        cell.alignment = Alignment(horizontal='left', vertical='center')

    for row_idx, row in enumerate(merged_df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx,
                    value='' if (value is None or (isinstance(value, float) and pd.isna(value))) else value)

    for col_idx, col_name in enumerate(final_cols, 1):
        col_letter = get_column_letter(col_idx)
        style      = column_styles.get(col_name, {})
        ws.column_dimensions[col_letter].width = style.get('width', max(len(col_name) + 4, 14))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return base64.b64encode(out.read()).decode('utf-8')


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    files       = request.files.getlist('files')
    all_styles  = {}
    file_data_list = []
    file_info   = []

    for f in files:
        if not f or not f.filename:
            continue
        ext = f.filename.rsplit('.', 1)[1].lower()
        if ext not in ('xlsx', 'xls', 'csv'):
            continue

        file_bytes = f.read()
        try:
            sheets_info   = read_all_sheets_info(file_bytes, f.filename)
            default_sheet = next(iter(sheets_info))

            file_data_list.append({
                'name': f.filename,
                'data': base64.b64encode(file_bytes).decode('utf-8'),
            })
            file_info.append({
                'name':          f.filename,
                'sheets':        sheets_info,
                'default_sheet': default_sheet,
            })

            styles = extract_header_styles(file_bytes, f.filename, sheet_name=default_sheet)
            for col, style in styles.items():
                if col not in all_styles:
                    all_styles[col] = style
        except Exception as e:
            return jsonify({'error': f'Error leyendo {f.filename}: {str(e)}'}), 400

    if not file_data_list:
        return jsonify({'error': 'Ningún archivo válido fue cargado'}), 400

    return jsonify({
        'column_styles': all_styles,
        'files':         file_info,
        'file_data':     file_data_list,
    })


@app.route('/preview', methods=['POST'])
def preview():
    data            = request.get_json()
    column_order    = data.get('columns', [])
    dedup_column    = data.get('dedup_column') or None
    priority_file   = data.get('priority_file') or None
    sheet_selection = data.get('sheet_selection', {})
    file_data_list  = data.get('file_data', [])

    if not column_order or not file_data_list:
        return jsonify({'error': 'Faltan parámetros'}), 400

    try:
        merged = _build_merged(file_data_list, column_order, sheet_selection, priority_file)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    final_cols = [c for c in column_order if c in merged.columns]

    similarity_suggestions = find_similar_values(merged[final_cols], final_cols)
    duplicate_groups       = find_duplicate_groups(merged, dedup_column)
    truncated              = len(duplicate_groups) == 50

    return jsonify({
        'columns':               final_cols,
        'total_rows':            len(merged),
        'duplicate_groups':      duplicate_groups,
        'similarity_suggestions': similarity_suggestions,
        'truncated_duplicates':  truncated,
    })


@app.route('/confirm', methods=['POST'])
def confirm():
    data            = request.get_json()
    column_order    = data.get('columns', [])
    column_styles   = data.get('column_styles', {})
    dedup_column    = data.get('dedup_column') or None
    priority_file   = data.get('priority_file') or None
    sheet_selection = data.get('sheet_selection', {})
    file_data_list  = data.get('file_data', [])
    homologation    = data.get('homologation', {})    # {col: {old: new}}
    rows_to_delete  = set(data.get('rows_to_delete', []))
    user_reviewed   = data.get('user_reviewed', False) # True = user saw & confirmed the review UI

    if not column_order or not file_data_list:
        return jsonify({'error': 'Faltan parámetros'}), 400

    try:
        merged = _build_merged(file_data_list, column_order, sheet_selection, priority_file)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    final_cols = [c for c in column_order if c in merged.columns]

    # Apply homologation
    for col, mapping in homologation.items():
        if col in merged.columns:
            merged[col] = merged[col].replace(mapping)

    before = len(merged)

    if user_reviewed:
        # User went through the review UI — only delete what they explicitly chose
        if rows_to_delete:
            merged = merged[~merged['__row_id__'].isin(rows_to_delete)]
        # If rows_to_delete is empty here it means the user chose "Conservar todos" → keep everything
    else:
        # Fast path (no duplicates found) → apply default dedup
        if dedup_column and dedup_column in merged.columns:
            merged = merged.drop_duplicates(subset=[dedup_column], keep='first')
        else:
            merged = merged.drop_duplicates(
                subset=[c for c in final_cols if c in merged.columns], keep='first')

    duplicates_removed = before - len(merged)
    merged = merged[final_cols]

    output_b64 = _write_excel(merged, final_cols, column_styles)

    return jsonify({
        'output_data':        output_b64,
        'total_rows':         len(merged),
        'duplicates_removed': duplicates_removed,
        'columns':            final_cols,
    })


@app.route('/merge', methods=['POST'])
def merge_files():
    """Fast path — kept for backward compatibility."""
    data            = request.get_json()
    column_order    = data.get('columns', [])
    column_styles   = data.get('column_styles', {})
    dedup_column    = data.get('dedup_column') or None
    priority_file   = data.get('priority_file') or None
    sheet_selection = data.get('sheet_selection', {})
    file_data_list  = data.get('file_data', [])

    if not column_order or not file_data_list:
        return jsonify({'error': 'Faltan parámetros'}), 400

    try:
        merged = _build_merged(file_data_list, column_order, sheet_selection, priority_file)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    final_cols = [c for c in column_order if c in merged.columns]
    merged     = merged[final_cols]

    before = len(merged)
    if dedup_column and dedup_column in merged.columns:
        merged = merged.drop_duplicates(subset=[dedup_column], keep='first')
    else:
        merged = merged.drop_duplicates(keep='first')

    output_b64 = _write_excel(merged, final_cols, column_styles)

    return jsonify({
        'output_data':        output_b64,
        'total_rows':         len(merged),
        'duplicates_removed': before - len(merged),
        'columns':            final_cols,
    })


if __name__ == '__main__':
    import os
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
