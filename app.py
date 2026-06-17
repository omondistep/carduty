import sqlite3
import os
import io
from flask import Flask, jsonify, request, render_template, g, send_file
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
DB_PATH = os.path.join(os.path.dirname(__file__), 'kraduty.db')
CURRENT_YEAR = 2025

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(exception):
    db = g.pop('db', None)
    if db is not None:
        db.close()

# Depreciation tables
DIRECT_DEPR = [
    (1, 2, 0.20), (2, 3, 0.30), (3, 4, 0.40), (4, 5, 0.50),
    (5, 6, 0.55), (6, 7, 0.60), (7, 8, 0.65),
]

PREV_REG_DEPR = [
    (1, 0.20), (2, 0.35), (3, 0.50), (4, 0.60), (5, 0.70),
    (6, 0.75), (7, 0.80), (8, 0.83), (9, 0.86), (10, 0.89),
    (11, 0.90), (12, 0.91), (13, 0.92), (14, 0.93), (15, 0.94),
    (16, 0.95),
]

def get_depreciation(years_old, is_direct_import=True):
    if is_direct_import:
        for lo, hi, rate in DIRECT_DEPR:
            if lo < years_old <= hi:
                return rate
        return 0.0
    else:
        for yr, rate in PREV_REG_DEPR:
            if years_old <= yr:
                return rate
        return 0.95

def calc_taxes(crsp, years_old, is_direct, vehicle_type, engine_cc, fuel):
    dep_rate = get_depreciation(years_old, is_direct)
    crsp_after_dep = crsp * (1 - dep_rate)

    if vehicle_type == 'motor_cycle':
        customs = (crsp_after_dep / 1.25) / 1.25 / 1.16
        import_duty = customs * 0.25
        excise_val = 0
        excise_duty = 12953.0
        vat_val = crsp_after_dep / 1.25 + excise_duty
        vat = vat_val * 0.16
        rdl = customs * 0.02
        idf = 0
        grand_total = import_duty + excise_duty + vat + rdl + idf
        components = {
            'customs_value': round(customs, 2),
            'import_duty': round(import_duty, 2),
            'excise_duty': round(excise_duty, 2),
            'vat_value': round(vat_val, 2),
            'vat': round(vat, 2),
            'rdl': round(rdl, 2),
            'idf': round(idf, 2),
            'grand_total': round(grand_total, 2),
            'excise_rate': 'Flat 12,953 KES',
            'import_duty_rate': '25%',
        }
        return components

    if vehicle_type == 'tractor':
        customs = (crsp_after_dep / 1.25) / 1.16
        import_duty = 0
        excise_duty = 0
        vat_val = crsp_after_dep / 1.25
        vat = vat_val * 0.16
        rdl = customs * 0.02
        idf = customs * 0.025
        grand_total = import_duty + excise_duty + vat + rdl + idf
        components = {
            'customs_value': round(customs, 2),
            'import_duty': round(import_duty, 2),
            'import_duty_rate': '0%',
            'excise_duty': round(excise_duty, 2),
            'excise_rate': '0%',
            'vat_value': round(vat_val, 2),
            'vat': round(vat, 2),
            'rdl': round(rdl, 2),
            'idf': round(idf, 2),
            'grand_total': round(grand_total, 2),
        }
        return components

    fuel_upper = fuel.upper() if fuel else ''
    is_electric = 'ELECTRIC' in fuel_upper and 'HYBRID' not in fuel_upper and 'PLUG' not in fuel_upper

    try:
        cc_str = engine_cc.split('(')[0].split(' kWh')[0].split(' ')[0] if engine_cc else '0'
        cc = float(cc_str) if cc_str.replace('.', '', 1).isdigit() else 0
    except (ValueError, TypeError):
        cc = 0

    if is_electric:
        actual_import_duty_rate = 0.25
        excise_rate = 0.10
    else:
        actual_import_duty_rate = 0.35
        if cc <= 1500:
            excise_rate = 0.20
        elif fuel_upper in ('GASOLINE', 'PETROL') and cc > 3000:
            excise_rate = 0.35
        elif fuel_upper in ('DIESEL') and cc > 2500:
            excise_rate = 0.35
        else:
            excise_rate = 0.25

    customs = (crsp_after_dep / 1.25) / 1.35 / (1 + excise_rate) / 1.16
    import_duty = customs * actual_import_duty_rate
    excise_val = customs + import_duty
    excise_duty = excise_val * excise_rate
    vat_val = customs + import_duty + excise_duty
    vat = vat_val * 0.16
    rdl = customs * 0.02
    idf = customs * 0.025
    grand_total = import_duty + excise_duty + vat + rdl + idf

    components = {
        'customs_value': round(customs, 2),
        'import_duty': round(import_duty, 2),
        'import_duty_rate': f'{int(actual_import_duty_rate*100)}%',
        'excise_duty': round(excise_duty, 2),
        'excise_rate': f'{int(excise_rate*100)}%',
        'vat_value': round(vat_val, 2),
        'vat': round(vat, 2),
        'rdl': round(rdl, 2),
        'idf': round(idf, 2),
        'grand_total': round(grand_total, 2),
    }
    return components

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/makes')
def api_makes():
    db = get_db()
    cur = db.execute("SELECT DISTINCT make FROM motor_vehicles WHERE make IS NOT NULL ORDER BY make")
    makes = [row['make'] for row in cur.fetchall()]
    return jsonify(makes)

@app.route('/api/models')
def api_models():
    make = request.args.get('make', '')
    db = get_db()
    cur = db.execute("SELECT DISTINCT model FROM motor_vehicles WHERE make = ? AND model IS NOT NULL ORDER BY model", (make,))
    models = [row['model'] for row in cur.fetchall()]
    return jsonify(models)

@app.route('/api/model_numbers')
def api_model_numbers():
    make = request.args.get('make', '')
    model = request.args.get('model', '')
    db = get_db()
    cur = db.execute(
        "SELECT DISTINCT model_number FROM motor_vehicles WHERE make = ? AND model = ? AND model_number IS NOT NULL ORDER BY model_number",
        (make, model))
    nums = [row['model_number'] for row in cur.fetchall()]
    return jsonify(nums)

@app.route('/api/variants')
def api_variants():
    make = request.args.get('make', '')
    model = request.args.get('model', '')
    model_no = request.args.get('model_number', '')
    db = get_db()
    if model_no:
        cur = db.execute(
            "SELECT DISTINCT transmission, drive_config, engine_capacity, body_type, fuel, crsp FROM motor_vehicles WHERE make = ? AND model = ? AND model_number = ?",
            (make, model, model_no))
    else:
        cur = db.execute(
            "SELECT DISTINCT transmission, drive_config, engine_capacity, body_type, fuel, crsp FROM motor_vehicles WHERE make = ? AND model = ?",
            (make, model))
    variants = []
    for row in cur.fetchall():
        variants.append(dict(zip(['transmission', 'drive_config', 'engine_capacity', 'body_type', 'fuel', 'crsp'], row)))
    return jsonify(variants)

@app.route('/api/mc_makes')
def api_mc_makes():
    db = get_db()
    cur = db.execute("SELECT DISTINCT make FROM motor_cycles WHERE make IS NOT NULL ORDER BY make")
    return jsonify([row['make'] for row in cur.fetchall()])

@app.route('/api/mc_models')
def api_mc_models():
    make = request.args.get('make', '')
    db = get_db()
    cur = db.execute("SELECT DISTINCT model FROM motor_cycles WHERE make = ? AND model IS NOT NULL ORDER BY model", (make,))
    return jsonify([row['model'] for row in cur.fetchall()])

@app.route('/api/mc_model_numbers')
def api_mc_model_numbers():
    make = request.args.get('make', '')
    model = request.args.get('model', '')
    db = get_db()
    cur = db.execute(
        "SELECT DISTINCT model_number FROM motor_cycles WHERE make = ? AND model = ? AND model_number IS NOT NULL ORDER BY model_number",
        (make, model))
    return jsonify([row['model_number'] for row in cur.fetchall()])

@app.route('/api/mc_variants')
def api_mc_variants():
    make = request.args.get('make', '')
    model = request.args.get('model', '')
    model_no = request.args.get('model_number', '')
    db = get_db()
    if model_no:
        cur = db.execute(
            "SELECT DISTINCT transmission, engine_capacity, fuel, crsp FROM motor_cycles WHERE make = ? AND model = ? AND model_number = ?",
            (make, model, model_no))
    else:
        cur = db.execute(
            "SELECT DISTINCT transmission, engine_capacity, fuel, crsp FROM motor_cycles WHERE make = ? AND model = ?",
            (make, model))
    variants = []
    for row in cur.fetchall():
        v = dict(zip(['transmission', 'engine_capacity', 'fuel', 'crsp'], row))
        if v['engine_capacity'] is not None:
            v['engine_capacity'] = str(int(v['engine_capacity']))
        v['drive_config'] = None
        v['body_type'] = None
        variants.append(v)
    return jsonify(variants)

@app.route('/api/tractor_makes')
def api_tractor_makes():
    db = get_db()
    cur = db.execute("SELECT DISTINCT make FROM tractors WHERE make IS NOT NULL ORDER BY make")
    return jsonify([row['make'] for row in cur.fetchall()])

@app.route('/api/tractor_models')
def api_tractor_models():
    make = request.args.get('make', '')
    db = get_db()
    cur = db.execute("SELECT model, horsepower, crsp FROM tractors WHERE make = ? AND model IS NOT NULL ORDER BY model", (make,))
    models = []
    for row in cur.fetchall():
        models.append({'model': row['model'], 'horsepower': row['horsepower'], 'crsp': row['crsp']})
    return jsonify(models)

@app.route('/api/calculate', methods=['POST'])
def api_calculate():
    data = request.get_json()
    crsp = float(data.get('crsp', 0))
    year_of_manufacture = int(data.get('yom', CURRENT_YEAR))
    month_of_manufacture = int(data.get('mom', 1))
    is_direct = data.get('is_direct', True)
    vehicle_type = data.get('vehicle_type', 'motor_vehicle')
    engine_cc = data.get('engine_capacity', '0')
    fuel = data.get('fuel', 'GASOLINE')

    if isinstance(is_direct, str):
        is_direct = is_direct == 'true'

    calendar_age = CURRENT_YEAR - year_of_manufacture
    if vehicle_type == 'motor_vehicle' and is_direct and calendar_age > 7:
        return jsonify({'error': 'Vehicles older than 7 years (manufactured before 2018) cannot be imported into Kenya.'}), 400

    current_month = 6
    total_months = (CURRENT_YEAR * 12 + current_month) - (year_of_manufacture * 12 + month_of_manufacture)
    age_years = total_months / 12.0

    result = calc_taxes(crsp, age_years, is_direct, vehicle_type, engine_cc, fuel)
    result['crsp'] = crsp
    result['yom'] = year_of_manufacture
    result['mom'] = month_of_manufacture
    result['age_years'] = round(age_years, 1)
    result['depreciation_rate'] = f'{get_depreciation(age_years, is_direct) * 100:.0f}%'

    return jsonify(result)

@app.route('/api/report/duties-below')
def api_report_duties_below():
    yom = int(request.args.get('yom', CURRENT_YEAR))
    mom = int(request.args.get('mom', 1))
    max_duty = float(request.args.get('max_duty', 500000))
    is_direct = request.args.get('is_direct', 'true') == 'true'
    vehicle_type = request.args.get('vehicle_type', 'motor_vehicle')

    db = get_db()

    if vehicle_type == 'motor_vehicle':
        cur = db.execute("SELECT * FROM motor_vehicles")
        rows = cur.fetchall()
    elif vehicle_type == 'motor_cycle':
        cur = db.execute("SELECT * FROM motor_cycles")
        rows = cur.fetchall()
    else:
        return jsonify({'error': 'Unsupported vehicle type'}), 400

    current_month = 6
    total_months = (CURRENT_YEAR * 12 + current_month) - (yom * 12 + mom)
    age_years = total_months / 12.0
    dep_rate = get_depreciation(age_years, is_direct)

    results = []
    for row in rows:
        row = dict(row)
        crsp = row.get('crsp')
        if not crsp:
            continue

        if vehicle_type == 'motor_vehicle' and is_direct and (CURRENT_YEAR - yom) > 7:
            continue

        engine_cc = str(row.get('engine_capacity') or '0')
        fuel = str(row.get('fuel') or 'GASOLINE')

        tax = calc_taxes(crsp, age_years, is_direct, vehicle_type, engine_cc, fuel)

        if tax['grand_total'] < max_duty:
            entry = {**row, **tax}
            entry['age_years'] = round(age_years, 1)
            entry['depreciation_rate'] = f'{dep_rate * 100:.0f}%'
            results.append(entry)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Duties Below Threshold"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2D6A4F", end_color="2D6A4F", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    if vehicle_type == 'motor_vehicle':
        cols = [
            'make', 'model', 'model_number', 'transmission', 'drive_config',
            'engine_capacity', 'body_type', 'gvw', 'seating', 'fuel', 'crsp',
            'age_years', 'depreciation_rate',
            'grand_total', 'customs_value', 'import_duty', 'excise_duty',
            'vat', 'rdl', 'idf'
        ]
        headers = [
            'Make', 'Model', 'Model Number', 'Transmission', 'Drive Config',
            'Engine Capacity', 'Body Type', 'GVW', 'Seating', 'Fuel', 'CRSP (KES)',
            'Age (yrs)', 'Depreciation Rate',
            'Total Duty (KES)', 'Customs Value (KES)', 'Import Duty (KES)',
            'Excise Duty (KES)', 'VAT (KES)', 'RDL (KES)', 'IDF (KES)'
        ]
    else:
        cols = [
            'make', 'model', 'model_number', 'transmission',
            'engine_capacity', 'seating', 'fuel', 'crsp',
            'age_years', 'depreciation_rate',
            'grand_total', 'customs_value', 'import_duty', 'excise_duty',
            'vat', 'rdl', 'idf'
        ]
        headers = [
            'Make', 'Model', 'Model Number', 'Transmission',
            'Engine Capacity', 'Seating', 'Fuel', 'CRSP (KES)',
            'Age (yrs)', 'Depreciation Rate',
            'Total Duty (KES)', 'Customs Value (KES)', 'Import Duty (KES)',
            'Excise Duty (KES)', 'VAT (KES)', 'RDL (KES)', 'IDF (KES)'
        ]

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    for row_idx, entry in enumerate(results, 2):
        for col_idx, key in enumerate(cols, 1):
            val = entry.get(key)
            if isinstance(val, float):
                val = round(val, 2)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right")

    for col_idx, _ in enumerate(headers, 1):
        ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else 'A'].bestFit = True
        ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else 'A'].width = max(12, len(headers[col_idx - 1]) + 4)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f'duties_below_{int(max_duty)}_{vehicle_type}_yom{yom}_mom{mom}.xlsx'
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5812))
    app.run(host='0.0.0.0', port=port, debug=False)
