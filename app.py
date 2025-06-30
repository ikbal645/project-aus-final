from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
import pandas as pd
import io

# --- Setup Flask ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'jadwal-sederhana'

# --- Setup Database SQLite ---
db_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'database.db')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- Model Jadwal ---
class Jadwal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    dosen = db.Column(db.String(100), nullable=False)
    matkul = db.Column(db.String(100), nullable=False)
    kelas = db.Column(db.String(20), nullable=False)
    hari = db.Column(db.String(20), nullable=False)
    jam_mulai = db.Column(db.String(5), nullable=False)
    jam_selesai = db.Column(db.String(5), nullable=False)
    ruangan = db.Column(db.String(50), nullable=False)

# --- Fungsi Bantu ---
def parse_time(t):
    try:
        return datetime.strptime(t, '%H:%M').time()
    except:
        return None

def bentrok(jam1_mulai, jam1_selesai, jam2_mulai, jam2_selesai):
    return jam1_mulai < jam2_selesai and jam2_mulai < jam1_selesai

def cek_bentrok(jadwal_baru):
    semua = Jadwal.query.filter_by(hari=jadwal_baru['hari']).all()
    jm, js = parse_time(jadwal_baru['jam_mulai']), parse_time(jadwal_baru['jam_selesai'])
    for j in semua:
        if j.kelas == jadwal_baru['kelas'] or j.dosen == jadwal_baru['dosen'] or j.ruangan == jadwal_baru['ruangan']:
            if bentrok(jm, js, parse_time(j.jam_mulai), parse_time(j.jam_selesai)):
                return True
    return False

# --- Rute Aplikasi ---
@app.route('/')
def index():
    return redirect('/jadwal')

@app.route('/jadwal')
def lihat_jadwal():
    semua = Jadwal.query.order_by(Jadwal.hari, Jadwal.jam_mulai).all()
    return render_template('jadwal.html', jadwal=semua)

@app.route('/tambah', methods=['GET', 'POST'])
def tambah():
    if request.method == 'POST':
        data = {
            'dosen': request.form['dosen'],
            'matkul': request.form['matkul'],
            'kelas': request.form['kelas'],
            'hari': request.form['hari'],
            'jam_mulai': request.form['jam_mulai'],
            'jam_selesai': request.form['jam_selesai'],
            'ruangan': request.form['ruangan']
        }
        if cek_bentrok(data):
            flash('❌ Jadwal bentrok dengan jadwal lain!', 'danger')
        else:
            j = Jadwal(**data)
            db.session.add(j)
            db.session.commit()
            flash('✅ Jadwal berhasil ditambahkan!', 'success')
        return redirect(url_for('lihat_jadwal'))
    return render_template('form.html')

@app.route('/hapus/<int:id>')
def hapus(id):
    j = Jadwal.query.get_or_404(id)
    db.session.delete(j)
    db.session.commit()
    flash('🗑️ Jadwal berhasil dihapus!', 'info')
    return redirect(url_for('lihat_jadwal'))

@app.route('/export')
def export_excel():
    semua = Jadwal.query.order_by(Jadwal.hari, Jadwal.jam_mulai).all()
    data = [{
        "Dosen": j.dosen,
        "Mata Kuliah": j.matkul,
        "Kelas": j.kelas,
        "Hari": j.hari,
        "Jam Mulai": j.jam_mulai,
        "Jam Selesai": j.jam_selesai,
        "Ruangan": j.ruangan
    } for j in semua]
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Jadwal')
    output.seek(0)

    return send_file(output,
                     download_name="jadwal.xlsx",
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# --- Jalankan Aplikasi ---
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
