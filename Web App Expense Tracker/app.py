from flask import Flask, render_template, request, send_file
import sqlite3
import openpyxl
from io import BytesIO
app = Flask(__name__)

# Create the database and table if not exists
def init_db():
    conn = sqlite3.connect('expenses.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            amount REAL
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date = request.form['date']
        category = request.form['category']
        amount = request.form['amount']

        conn = sqlite3.connect('expenses.db')
        c = conn.cursor()
        c.execute("INSERT INTO expenses (date, category, amount) VALUES (?, ?, ?)",
                  (date, category, amount))
        conn.commit()
        conn.close()

    # Fetch all expenses from DB
    conn = sqlite3.connect('expenses.db')
    c = conn.cursor()
    c.execute("SELECT date, category, amount FROM expenses ORDER BY date DESC")
    expenses = c.fetchall()
    conn.close()

    return render_template('index.html', expenses=expenses)


if __name__ == '__main__':
    init_db()
    app.run(debug=True)


from openpyxl import Workbook
from flask import send_file
import io

@app.route('/export')
def export_excel():
    conn = sqlite3.connect('expenses.db')
    c = conn.cursor()
    c.execute('SELECT date, category, amount FROM expenses ORDER BY date DESC')
    data = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.append(['Date', 'Category', 'Amount'])

    for row in data:
        ws.append(row)

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(file_stream, as_attachment=True, download_name='expenses.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/export')
def export():
    conn = sqlite3.connect('expenses.db')
    c = conn.cursor()
    c.execute("SELECT date, category, amount FROM expenses ORDER BY date DESC")
    data = c.fetchall()
    conn.close()

    # Create Excel file
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expenses"
    ws.append(["Date", "Category", "Amount"])

    for row in data:
        ws.append(row)

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, download_name="expenses.xlsx", as_attachment=True)
