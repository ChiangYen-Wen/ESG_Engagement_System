from flask import Flask, render_template, request, redirect
import pyodbc
import os
from datetime import datetime

app = Flask(__name__)

# 資料庫路徑設定（請確保 ESG_DB.accdb 與本檔案在同一資料夾）
db_path = os.path.abspath("ESG_DB.accdb")
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    f'DBQ={db_path};'
)

# 取得下拉選單選項（唯一值）
def get_unique_values(column_name):
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(f"SELECT DISTINCT {column_name} FROM ESG_report WHERE {column_name} IS NOT NULL")
        return [row[0] for row in cursor.fetchall()]

# 🔍 查詢主頁
@app.route('/', methods=['GET', 'POST'])
def index():
    results = []
    stock_ids = get_unique_values("stock_id")
    analysts = get_unique_values("analyst")
    purposes = get_unique_values("purpose")
    success = request.args.get("success")


    if request.method == 'POST':
        stock_id = request.form.get('stock_id')
        analyst = request.form.get('analyst')
        purpose = request.form.get('purpose')
        date_start = request.form.get('date_start')
        date_end = request.form.get('date_end')

        sql = "SELECT stock_id, stock_name, analyst, purpose, ESG_category FROM ESG_report WHERE 1=1"
        params = []

        if stock_id:
            sql += " AND stock_id = ?"
            params.append(stock_id)
        if analyst:
            sql += " AND analyst = ?"
            params.append(analyst)
        if purpose:
            sql += " AND purpose = ?"
            params.append(purpose)
        if date_start:
            sql += " AND date >= ?"
            params.append(date_start)
        if date_end:
            sql += " AND date <= ?"
            params.append(date_end)

        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql, params)
            results = cursor.fetchall()

    return render_template(
        'index.html',
        results=results,
        stock_ids=stock_ids,
        analysts=analysts,
        purposes=purposes,
        success=success
    )

# 📝 新增報告功能（重點加強日期時間轉換）
@app.route('/add', methods=['POST'])
def add():
    form = request.form

    def parse_date(val):
        try:
            return datetime.strptime(val, '%Y-%m-%d').date() if val else None
        except:
            return None

    def parse_datetime(val):
        try:
            return datetime.strptime(val, '%Y-%m-%dT%H:%M') if val else None
        except:
            return None

    values = [
        parse_date(form.get("date")),
        parse_datetime(form.get("upload_time")),
        form.get("stock_id"),
        form.get("stock_name"),
        form.get("analyst"),
        form.get("purpose"),
        form.get("ESG_category"),
        form.get("shareholders_meeting"),
        form.get("question"),
        form.get("answer"),
        form.get("files_type"),
        form.get("remark"),
    ]

    sql = """
        INSERT INTO ESG_report 
        ([date], [upload_time], [stock_id], [stock_name], [analyst], [purpose], [ESG_category],
         [shareholders_meeting], [question], [answer], [files_type], [remark])
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        conn.commit()

    return redirect('/?success=add')

# 🔎 查看單筆報告詳細資料
@app.route('/detail')
def detail():
    stock_id = request.args.get('stock_id')
    stock_name = request.args.get('stock_name')
    analyst = request.args.get('analyst')

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM ESG_report 
            WHERE stock_id = ? AND stock_name = ? AND analyst = ?
        """, (stock_id, stock_name, analyst))
        row = cursor.fetchone()
        report = dict(zip([col[0] for col in cursor.description], row)) if row else None

    return render_template("detail.html", report=report)

@app.route('/edit', methods=['GET', 'POST'])
def edit():
    stock_id = request.args.get('stock_id')
    stock_name = request.args.get('stock_name')
    analyst = request.args.get('analyst')

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM ESG_report 
            WHERE stock_id = ? AND stock_name = ? AND analyst = ?
        """, (stock_id, stock_name, analyst))
        row = cursor.fetchone()
        report = dict(zip([col[0] for col in cursor.description], row)) if row else None

    return render_template("edit.html", report=report)

# 💾 更新報告
@app.route('/update', methods=['POST'])
def update():
    form = request.form

    def parse_date(val):
        try:
            return datetime.strptime(val, '%Y-%m-%d').date() if val else None
        except:
            return None

    def parse_datetime(val):
        try:
            return datetime.strptime(val, '%Y-%m-%dT%H:%M') if val else None
        except:
            return None

    values = [
        parse_date(form.get("date")),
        parse_datetime(form.get("upload_time")),
        form.get("stock_name"),
        form.get("analyst"),
        form.get("purpose"),
        form.get("ESG_category"),
        form.get("shareholders_meeting"),
        form.get("question"),
        form.get("answer"),
        form.get("files_type"),
        form.get("remark"),
        form.get("stock_id"),  # WHERE 條件
    ]

    sql = """
        UPDATE ESG_report SET
            [date] = ?, [upload_time] = ?, [stock_name] = ?, [analyst] = ?, [purpose] = ?,
            [ESG_category] = ?, [shareholders_meeting] = ?, [question] = ?, [answer] = ?,
            [files_type] = ?, [remark] = ?
        WHERE [stock_id] = ?
    """

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        conn.commit()

    return redirect('/?success=update')

@app.route('/delete', methods=['POST'])
def delete():
    stock_id = request.form.get('stock_id')
    stock_name = request.form.get('stock_name')
    analyst = request.form.get('analyst')

    sql = """
        DELETE FROM ESG_report 
        WHERE stock_id = ? AND stock_name = ? AND analyst = ?
    """

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, (stock_id, stock_name, analyst))
        conn.commit()

    return redirect('/?success=delete')



# ✅ 啟動 Flask Server
if __name__ == '__main__':
    import webbrowser
    webbrowser.open("http://127.0.0.1:5000")
    app.run(debug=True)

