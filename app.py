import os
import sqlite3
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
import nepali_datetime
from datetime import datetime
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'secret_key_for_session'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ------------------ USER DATABASE ------------------

def init_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    password TEXT NOT NULL,
                    email TEXT,
                    phone_number TEXT
                )''')
    conn.commit()
    conn.close()

init_db()

# ------------------ EXPORT USERS TO EXCEL ------------------

def export_users_to_excel():
    excel_path = 'user_details.xlsx'
    if os.path.exists(excel_path):
        os.remove(excel_path)
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query("SELECT username, email, phone_number FROM users", conn)
    conn.close()
    df.to_excel(excel_path, index=False)

# ------------------ HELPERS ------------------

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def bs_to_ad(value):
    try:
        if pd.isna(value):
            return None
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")

        bs_date = str(value).replace("-", "/").strip()
        parts = bs_date.split("/")

        if len(parts) != 3:
            return None

        if int(parts[0]) > 2000:
            y, m, d = map(int, parts)
        else:
            d, m, y = map(int, parts)

        bs_obj = nepali_datetime.date(y, m, d)
        ad_date = bs_obj.to_datetime_date()
        return ad_date.strftime("%d/%m/%Y")
    except:
        return None

# ------------------ AUTH ROUTES ------------------

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        password = generate_password_hash(request.form['password'])
        email = request.form['email']
        phone = request.form['phone']

        try:
            conn = sqlite3.connect('users.db')
            c = conn.cursor()
            c.execute("INSERT INTO users (username, password, email, phone_number) VALUES (?, ?, ?, ?)",
                      (username, password, email, phone))
            conn.commit()
            conn.close()

            # Update Excel file on new registration
            export_users_to_excel()

            flash('Registration successful. Please log in.')
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash('Username already exists.')
            return redirect(url_for('register'))

    return render_template('register.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username = ?", (username,))
        user = c.fetchone()
        conn.close()

        if user and check_password_hash(user[2], password):
            session['user'] = username
            return redirect(url_for('index'))
        else:
            flash('Invalid username or password.')
            return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('user', None)
    flash('You have been logged out.')
    return redirect(url_for('login'))

# ------------------ MAIN FUNCTIONALITY ------------------

@app.route('/', methods=['GET', 'POST'])
def index():
    if 'user' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        file = request.files.get('file')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            df = pd.read_excel(filepath)

            # Try to detect BS date column
            date_col = None
            for col in df.columns:
                for val in df[col]:
                    try:
                        if pd.notna(val):
                            s = str(val).replace("-", "/")
                            parts = s.split("/")
                            if len(parts) == 3 and (int(parts[0]) >= 2000 or int(parts[2]) >= 2000):
                                date_col = col
                                break
                    except:
                        continue
                if date_col:
                    break

            if not date_col:
                return render_template('index.html', columns=df.columns, file=filename,
                                       needs_column_selection=True, username=session['user'])

            # Auto convert and insert to right of date column
            converted = df[date_col].apply(bs_to_ad)
            col_index = df.columns.get_loc(date_col)
            df.insert(col_index + 1, "English_Date", converted)

            output_path = os.path.join(app.config['UPLOAD_FOLDER'], "converted_" + filename)
            df.to_excel(output_path, index=False)
            return send_file(output_path, as_attachment=True)

        flash("Invalid file format. Please upload an Excel file.")
        return redirect(url_for('index'))

    return render_template('index.html', username=session['user'])

@app.route('/convert', methods=['POST'])
def convert_with_column():
    if 'user' not in session:
        return redirect(url_for('login'))

    filename = request.form['filename']
    column = request.form['column']
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    try:
        df = pd.read_excel(filepath)

        converted = df[column].apply(bs_to_ad)
        col_index = df.columns.get_loc(column)
        df.insert(col_index + 1, "English_Date", converted)

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], "converted_" + filename)
        df.to_excel(output_path, index=False)
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        flash("Error during conversion: " + str(e))
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
