from flask import Flask, render_template, request, jsonify, send_file
import sqlite3
from datetime import datetime, timedelta
import calendar
import pandas as pd
from io import BytesIO
import os
import json

app = Flask(__name__)

# Database setup (same as before)
DB_NAME = "daily_reports_advanced.db"

BRANCHES = [
    "MAIN", "ANNEX", "CONCEPCION", "LEGAZPI", "TABUCO",
    "TABUC ANNEX", "PILI", "DAET", "PILI NEW", "DIVERSION", "LEGAZPI ANNEX"
]

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS daily_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            branch TEXT NOT NULL,
            date TEXT NOT NULL,
            sales_am REAL DEFAULT 0,
            sales_pm REAL DEFAULT 0,
            rooms_am INTEGER DEFAULT 0,
            rooms_pm INTEGER DEFAULT 0,
            water_reading REAL DEFAULT 0,
            electricity_reading REAL DEFAULT 0,
            UNIQUE(branch, date)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# Routes
@app.route('/')
def index():
    return render_template('index.html', branches=BRANCHES)

@app.route('/api/dashboard', methods=['POST'])
def dashboard():
    data = request.json
    from_date = data.get('from_date')
    to_date = data.get('to_date')
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # Get summary data
    c.execute('''
        SELECT 
            COALESCE(SUM(sales_am + sales_pm), 0),
            COALESCE(SUM(rooms_am + rooms_pm), 0),
            COALESCE(SUM(water_reading), 0),
            COALESCE(SUM(electricity_reading), 0)
        FROM daily_reports 
        WHERE date >= ? AND date <= ?
    ''', (from_date, to_date))
    
    totals = c.fetchone()
    
    # Get branch-wise data
    branch_data = []
    for branch in BRANCHES:
        c.execute('''
            SELECT 
                COALESCE(SUM(sales_am + sales_pm), 0),
                COALESCE(SUM(rooms_am + rooms_pm), 0),
                COUNT(*) as days
            FROM daily_reports 
            WHERE branch = ? AND date >= ? AND date <= ?
        ''', (branch, from_date, to_date))
        row = c.fetchone()
        branch_data.append({
            'branch': branch,
            'sales': row[0],
            'rooms': row[1],
            'days': row[2]
        })
    
    conn.close()
    
    return jsonify({
        'totals': {
            'sales': totals[0],
            'rooms': totals[1],
            'water': totals[2],
            'electricity': totals[3]
        },
        'branches': branch_data
    })

@app.route('/api/save_data', methods=['POST'])
def save_data():
    data = request.json
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    try:
        c.execute('''
            INSERT OR REPLACE INTO daily_reports
            (branch, date, sales_am, sales_pm, rooms_am, rooms_pm, water_reading, electricity_reading)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['branch'], data['date'],
            float(data['sales_am']), float(data['sales_pm']),
            int(data['rooms_am']), int(data['rooms_pm']),
            float(data['water_reading']), float(data['electricity_reading'])
        ))
        conn.commit()
        return jsonify({'success': True, 'message': 'Data saved successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})
    finally:
        conn.close()

@app.route('/api/daily_report', methods=['POST'])
def daily_report():
    data = request.json
    from_date = data.get('from_date')
    to_date = data.get('to_date')
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    c.execute('''
        SELECT date, branch, sales_am, sales_pm, rooms_am, rooms_pm, water_reading, electricity_reading
        FROM daily_reports 
        WHERE date >= ? AND date <= ?
        ORDER BY date, branch
    ''', (from_date, to_date))
    
    rows = c.fetchall()
    
    # Calculate consumption for each row
    result = []
    for row in rows:
        date, branch, sales_am, sales_pm, rooms_am, rooms_pm, water, elec = row
        
        # Get previous readings
        c.execute('''
            SELECT water_reading, electricity_reading FROM daily_reports
            WHERE branch = ? AND date < ?
            ORDER BY date DESC LIMIT 1
        ''', (branch, date))
        prev = c.fetchone()
        
        water_cons = water - prev[0] if prev and water >= prev[0] else 0
        elec_cons = elec - prev[1] if prev and elec >= prev[1] else 0
        
        result.append({
            'date': date,
            'branch': branch,
            'sales_am': sales_am,
            'sales_pm': sales_pm,
            'total_sales': sales_am + sales_pm,
            'rooms_am': rooms_am,
            'rooms_pm': rooms_pm,
            'total_rooms': rooms_am + rooms_pm,
            'water_cons': water_cons,
            'elec_cons': elec_cons
        })
    
    conn.close()
    return jsonify(result)

@app.route('/api/export/<report_type>', methods=['POST'])
def export_report(report_type):
    data = request.json
    
    if report_type == 'daily':
        # Create DataFrame from data
        df = pd.DataFrame(data['data'])
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Daily Report')
        
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'daily_report_{datetime.now().strftime("%Y%m%d")}.xlsx'
        )

if __name__ == '__main__':
    app.run(debug=True)