import sqlite3
from flask import Flask, render_template, request, redirect, url_for, session, flash

app = Flask(__name__)
app.secret_key = 'alhamdi_tech_secret_key'  # مفتاح تشفير الجلسات

# دالة الاتصال بقاعدة البيانات
def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

# دالة إنشائية لتهيئة قاعدة البيانات وجداولها
def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # جدول الخدمات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            category TEXT NOT NULL,
            description TEXT NOT NULL,
            image TEXT NOT NULL
        )
    ''')
    
    # جدول طلبات العملاء
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            service_type TEXT NOT NULL,
            details TEXT
        )
    ''')

    # جدول المستخدمين (لـ Admin)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    ''')
    
    # إضافة بيانات أولية للخدمات إذا كان الجدول فارغاً
    cursor.execute("SELECT COUNT(*) FROM services")
    if cursor.fetchone()[0] == 0:
        default_services = [
            ("أنظمة الكاميرات والمراقبة", "كاميرات", "تركيب وبرمجة أجهزة NVR/DVR وأنظمة IP مع إمكانية المراقبة عن بعد.", "cctv.jpg"),
            ("إعداد الشبكات والربط اللاسلكي", "شبكات", "ربط الفروع، ربط P2P اللاسلكي، وتكوين سويتشات وسيرفرات Cisco/MikroTik.", "network.jpg"),
            ("الصيانة والدعم الفني", "صيانة", "فحص وصيانة أجهزة الكمبيوتر، السيرفرات، والأجهزة المحمولة وتحديث الأنظمة.", "maintenance.jpg")
        ]
        cursor.executemany("INSERT INTO services (title, category, description, image) VALUES (?, ?, ?, ?)", default_services)
        
    # إضافة حساب مدير افتراضي إذا لم يكن موجوداً
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", ("admin", "Yemen@123"))

    conn.commit()
    conn.close()

# تهيئة قاعدة البيانات عند تشغيل التطبيق
init_db()

# --- المسارات (Routes) ---

# الصفحة الرئيسية
@app.route('/')
def home():
    conn = get_db_connection()
    services = conn.execute("SELECT * FROM services LIMIT 3").fetchall()
    conn.close()
    return render_template('index.html', services=services)

# صفحة الخدمات
@app.route('/services')
def services():
    conn = get_db_connection()
    all_services = conn.execute("SELECT * FROM services").fetchall()
    conn.close()
    return render_template('services.html', services=all_services)

# صفحة من نحن
@app.route('/about')
def about():
    return render_template('about.html')

# صفحة طلب خدمة / التواصل
@app.route('/contact', methods=['GET', 'POST'])
def contact():
    if request.method == 'POST':
        client_name = request.form.get('name')
        phone = request.form.get('phone')
        service_type = request.form.get('service_type')
        details = request.form.get('details')

        if client_name and phone and service_type:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO requests (client_name, phone, service_type, details) VALUES (?, ?, ?, ?)",
                (client_name, phone, service_type, details)
            )
            conn.commit()
            conn.close()
            flash("تم إرسال طلبك بنجاح! سنقوم بالتواصل معك قريباً.", "success")
            return redirect(url_for('contact'))
        else:
            flash("يرجى ملء جميع الحقول المطلوبة.", "danger")

    return render_template('contact.html')

# صفحة تسجيل الدخول
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        conn = get_db_connection()
        user = conn.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password)).fetchone()
        conn.close()

        if user:
            session['logged_in'] = True
            session['username'] = username
            flash("تم تسجيل الدخول بنجاح!", "success")
            return redirect(url_for('dashboard'))
        else:
            flash("اسم المستخدم أو كلمة المرور غير صحيحة.", "danger")

    return render_template('login.html')

# لوحة التحكم الخاصة بالمدير (Dashboard)
@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    conn = get_db_connection()
    requests_list = conn.execute("SELECT * FROM requests ORDER BY id DESC").fetchall()
    conn.close()
    return render_template('dashboard.html', requests=requests_list)

# تسجيل الخروج
@app.route('/logout')
def logout():
    session.clear()
    flash("تم تسجيل الخروج بنجاح.", "info")
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)