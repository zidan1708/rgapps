from flask import Flask, request, jsonify, render_template
import os
import pandas as pd
from sqlalchemy import create_engine, inspect
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Konfigurasi
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
DATABASE_URI = 'sqlite:///database.db'  # Ganti dengan URI database Anda

# Buat folder upload jika belum ada
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Fungsi untuk memeriksa ekstensi file
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Periksa apakah file ada dalam request
    if 'excelFile' not in request.files:
        return jsonify({'success': False, 'message': 'Tidak ada file yang dipilih'})
    
    file = request.files['excelFile']
    table_name = request.form.get('tableName', '').strip()
    
    # Validasi input
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Tidak ada file yang dipilih'})
    
    if not table_name:
        return jsonify({'success': False, 'message': 'Nama tabel tidak boleh kosong'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'message': 'Format file tidak didukung. Harap upload file Excel (.xlsx atau .xls)'})
    
    try:
        # Simpan file yang diupload
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Baca file Excel
        if filename.endswith('.xlsx'):
            df = pd.read_excel(filepath, engine='openpyxl')
        else:
            df = pd.read_excel(filepath)
        
        # Bersihkan nama kolom (hapus spasi dan karakter khusus)
        df.columns = [col.strip().replace(' ', '_').lower() for col in df.columns]
        
        # Simpan ke database
        engine = create_engine(DATABASE_URI)
        
        # Periksa apakah tabel sudah ada
        inspector = inspect(engine)
        if table_name in inspector.get_table_names():
            # Jika tabel sudah ada, tambahkan data baru
            df.to_sql(table_name, engine, if_exists='append', index=False)
            operation = 'Data berhasil ditambahkan ke tabel yang sudah ada'
        else:
            # Jika tabel belum ada, buat tabel baru
            df.to_sql(table_name, engine, index=False)
            operation = 'Tabel baru berhasil dibuat dan data dimasukkan'
        
        # Dapatkan informasi tentang data yang diimpor
        row_count = len(df)
        columns = list(df.columns)
        
        # Hapus file setelah diproses
        os.remove(filepath)
        
        return jsonify({
            'success': True,
            'message': operation,
            'table_name': table_name,
            'row_count': row_count,
            'columns': columns
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})

if __name__ == '__main__':
    app.run(debug=True)
