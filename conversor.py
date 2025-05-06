from flask import Blueprint, render_template, request, flash, send_file, url_for, current_app, session, redirect
import pandas as pd
import os
from datetime import timedelta

conversor_bp = Blueprint('conversor', __name__, template_folder='templates')

DAY_NAMES = {
    0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves',
    4: 'Viernes', 5: 'Sábado', 6: 'Domingo'
}

SERVICES = [
    'Sop_Conectividad', 'Sop_Flow', 'Esp_CATV', 'Esp_Movil', 'Esp_XDSL',
    'Digital', 'CBS', 'SMB_TecnicaIN', 'SMB_Digital'
]

@conversor_bp.route('/conversor', methods=['GET', 'POST'])
def conversor():
    download_url = None
    if request.method == 'POST':
        servicio = request.form.get('servicio')
        prog_file = request.files.get('prog_file')
        if not servicio or not prog_file:
            flash('Selecciona un servicio y un archivo.', 'warning')
            return render_template('conversor.html', title='Conversor', services=SERVICES, download_url=None)

        upload_dir = current_app.config.get('UPLOAD_FOLDER', os.getcwd())
        in_path = os.path.join(upload_dir, 'input_prog.xlsx')
        prog_file.save(in_path)

        df = pd.read_excel(in_path)
        if 'SERVICIO' in df.columns:
            key = servicio.split('_')[-1]
            df = df[df['SERVICIO'].str.contains(key, case=False, na=False)]

        df = df.dropna(subset=['Nombres_Presentes'])
        records = []
        for _, row in df.iterrows():
            names = [n.strip() for n in str(row['Nombres_Presentes']).split(';')]
            for name in names:
                if not name:
                    continue
                if ',' in name:
                    parts = [p.strip() for p in name.split(',')]
                    name = f"{parts[1]} {parts[0]}"
                new_row = row.copy()
                new_row['Nombre'] = name.upper()
                records.append(new_row)
        df = pd.DataFrame(records)

        nomina_path = session.get('nomina_path')
        if not nomina_path or not os.path.exists(nomina_path):
            flash("No se encontró la nómina cargada.", 'danger')
            return render_template('conversor.html', title='Conversor', services=SERVICES)

        df_nom = pd.read_excel(nomina_path)
        df_nom.columns = df_nom.columns.str.strip()
        df_nom = df_nom[['NOMBRE', 'NUEVO SUPERIOR']].rename(columns={'NOMBRE': 'Nombre'})
        df_nom['Nombre'] = df_nom['Nombre'].str.strip().str.upper()

        df = df.merge(df_nom, on='Nombre', how='left')

        if 'NUEVO SUPERIOR' not in df.columns or df['NUEVO SUPERIOR'].isnull().all():
            flash("No se pudieron asignar líderes. Revisá que los nombres coincidan.", 'danger')
            return render_template('conversor.html', title='Conversor', services=SERVICES)

        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df = df.dropna(subset=['Fecha'])
        df['Fecha'] = df['Fecha'].dt.date
        df['Intervalo_dt'] = pd.to_datetime(df['Intervalo'], format='%H:%M').dt.time
        df['Intervalo'] = df['Intervalo_dt'].astype(str)
        df['Semana'] = df['Fecha'].apply(lambda d: d - timedelta(days=d.weekday()))
        df = df.sort_values(by=['Fecha', 'Intervalo_dt'])

        file_name = f'convertido_tabs_{servicio}.xlsx'
        out_path = os.path.join(upload_dir, file_name)

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            for semana, group in df.groupby('Semana'):
                hoja = semana.strftime('Sem %Y-%m-%d')
                tmp = group.copy()
                tmp['Dia_Num'] = pd.to_datetime(tmp['Fecha']).apply(lambda d: d.weekday())
                tmp['Presente'] = 1
                tmp = tmp.sort_values(by=['Intervalo_dt', 'Nombre'])

                # Presencia por día
                pivot = tmp.pivot_table(
                    index='Nombre', columns='Dia_Num',
                    values='Presente', aggfunc='sum', fill_value=0
                )
                for i in range(7):
                    if i not in pivot.columns:
                        pivot[i] = 0
                pivot = pivot.reindex(columns=range(7), fill_value=0)
                pivot.rename(columns=DAY_NAMES, inplace=True)
                for col in DAY_NAMES.values():
                    pivot[col] = (pivot[col] > 0).astype(int)
                pivot.reset_index(inplace=True)

                # Líder e intervalo
                li_map = tmp.groupby('Nombre').agg({'NUEVO SUPERIOR': 'first', 'Intervalo': 'first'}).reset_index()
                pivot = pivot.merge(li_map, on='Nombre', how='left')

                # Asignación distribuida de breaks
                break_cols = []
                for day_num, day_name in DAY_NAMES.items():
                    break_col = f'Break_{day_name}'
                    break_cols.append(break_col)

                    day_data = tmp[tmp['Dia_Num'] == day_num]
                    grouped = day_data.groupby('Nombre')

                    person_break_windows = {}
                    for name, group_day in grouped:
                        horas = sorted([pd.to_datetime(h).time() for h in group_day['Intervalo']])
                        if len(horas) < 3:
                            continue
                        inicio = pd.to_datetime(horas[0].strftime('%H:%M'))
                        fin = pd.to_datetime(horas[-1].strftime('%H:%M'))
                        ventana_inicio = inicio + timedelta(hours=2)
                        ventana_fin = fin - timedelta(hours=2)
                        posibles = [pd.to_datetime(h.strftime('%H:%M')) for h in horas if ventana_inicio <= pd.to_datetime(h.strftime('%H:%M')) <= ventana_fin]
                        if posibles:
                            person_break_windows[name] = posibles

                    used_slots = set()
                    break_map = {}
                    for i, (name, posibles) in enumerate(sorted(person_break_windows.items())):
                        for candidato in posibles:
                            if candidato not in used_slots:
                                break_map[name] = candidato.strftime('%H:%M')
                                used_slots.add(candidato)
                                break
                        else:
                            elegido = posibles[len(posibles) // 2]
                            break_map[name] = elegido.strftime('%H:%M')

                    pivot[break_col] = pivot['Nombre'].map(break_map).fillna("")

                cols = ['Nombre', 'NUEVO SUPERIOR', 'Intervalo'] + list(DAY_NAMES.values()) + break_cols
                pivot = pivot[cols]
                pivot.to_excel(writer, sheet_name=hoja, index=False)

        session['last_file'] = file_name
        download_url = url_for('conversor.download')

    return render_template('conversor.html', title='Conversor', services=SERVICES, download_url=download_url)

@conversor_bp.route('/conversor/download')
def download():
    upload_dir = current_app.config.get('UPLOAD_FOLDER', os.getcwd())

    last_file = session.get('last_file')
    if last_file and os.path.exists(os.path.join(upload_dir, last_file)):
        path = os.path.join(upload_dir, last_file)
    else:
        files = [f for f in os.listdir(upload_dir) if f.startswith("convertido_tabs_") and f.endswith(".xlsx")]
        if not files:
            flash("No se encontró el archivo para descargar.", "danger")
            return redirect(url_for('conversor.conversor'))
        files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(upload_dir, f)), reverse=True)
        path = os.path.join(upload_dir, files[0])

    return send_file(path, as_attachment=True, download_name=os.path.basename(path))
