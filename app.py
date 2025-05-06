from flask import Flask, redirect, url_for, render_template, request, session, flash
import os

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = "cambia_esto_por_una_clave_segura"
app.config['UPLOAD_FOLDER'] = os.path.abspath(os.path.dirname(__file__))

# Registrar blueprints
from layout import layout_bp
from conversor import conversor_bp
from programacion import programacion_bp

app.register_blueprint(layout_bp)
app.register_blueprint(conversor_bp)
app.register_blueprint(programacion_bp)

# Credenciales válidas
default_credentials = {"Workforce": "Flow2025"}

@app.route('/', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        user = request.form.get('username')
        pwd  = request.form.get('password')
        if default_credentials.get(user) == pwd:
            session.clear()
            session['logged_in'] = True
            # Una vez logueado, primero subimos nómina
            return redirect(url_for('upload_nomina'))
        else:
            flash('Credenciales incorrectas', 'warning')
    return render_template('login.html', title='Login')

@app.route('/nomina', methods=['GET','POST'])
def upload_nomina():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    if request.method == 'POST':
        nomina_file = request.files.get('nomina')
        if not nomina_file:
            flash('Selecciona un archivo de nómina (.xlsx)', 'warning')
        else:
            path = os.path.join(app.config['UPLOAD_FOLDER'], 'nomina.xlsx')
            nomina_file.save(path)
            session['nomina_path'] = path
            return redirect(url_for('menu'))
    return render_template('nomina.html', title='Carga de Nómina')

@app.route('/menu')
def menu():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    if 'nomina_path' not in session:
        return redirect(url_for('upload_nomina'))
    return render_template('index.html', title='Menú Principal')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
