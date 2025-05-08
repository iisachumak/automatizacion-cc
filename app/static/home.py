from flask import Blueprint, render_template, request, flash, redirect, url_for, send_from_directory
import pandas as pd
import os
from werkzeug.utils import secure_filename
import uuid

home_bp = Blueprint('home', __name__)

# Configuración para archivos temporales
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'app', 'static', 'descargas')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@home_bp.route('/')
def home():
    return render_template('index.html')

@home_bp.route('/static/descargas/<filename>')#static\downloads\financiaciones_f260d1d8043241da8b7524805c2818f2.xlsx
def download_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

@home_bp.route('/automatizacion', methods=['POST'])
def automaticacion():
    if 'excel' not in request.files:
        flash('No se encontró ningún archivo', 'danger')
        return redirect(url_for('home.home'))
    
    excel_file = request.files['excel']
    print('hay archivo')

    if excel_file.filename == '':
        flash('Archivo vacío', 'warning')
        return redirect(url_for('home.home'))

    try:
        df = pd.read_excel(excel_file, header=[0, 1])
        print(df.columns.tolist())

        resultados = []

        for index, row in df.iterrows():
            #SKU
            sku = row[('Unnamed: 0_level_0', 'CODIGO')]
            
            #CS
            cs_3 = "3 CS " if pd.notna(row[('CUOTA SIMPLE', 3)]) else ''
            cs_6 = "6 CS " if pd.notna(row[('CUOTA SIMPLE', 6)]) else ''
            cs_1 = "1 CS " if pd.isna(row[('CUOTA SIMPLE', 3)]) and pd.isna(row[('CUOTA SIMPLE', 6)]) else ''

            #VMP
            vmp_18 = "18 VMP " if pd.notna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 18)]) else ''
            vmp_15 = "15 VMP " if pd.notna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 15)]) else ''
            vmp_12 = "12 VMP " if pd.notna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 12)]) else ''
            vmp_9 = "9 VMP " if pd.notna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 9)]) else ''
            vmp_1 = "1 VMP " if pd.isna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 18)]) and pd.isna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 15)]) and pd.isna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 12)]) and pd.isna(row[('VISA /MASTER (Adheridos PAYWAY) (Bancarizadas NO Santander, Macro, Galicia, Nación y Frances)', 9)]) else ''

            #VMNP
            vmnp_18 = "18 VMNP " if pd.notna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 18)]) else ''
            vmnp_15 = "15 VMNP " if pd.notna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 15)]) else ''
            vmnp_12 = "12 VMNP " if pd.notna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 12)]) else ''
            vmnp_9 = "9 VMNP " if pd.notna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 9)]) else ''
            vmnp_1 = "1 VMNP " if pd.isna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 18)]) and pd.isna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 15)]) and pd.isna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 12)]) and pd.isna(row[('VISA /MASTER (NO Adheridos PAYWAY) Santander, Macro, Galicia, Nación y Frances', 9)]) else ''
            
            #AM
            am_12 = "12 AM " if pd.notna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 12)]) else ''
            am_9 = "9 AM " if pd.notna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 9)]) else ''
            am_6 = "6 AM " if pd.notna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 6)]) else ''
            am_3 = "3 AM " if pd.notna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 3)]) else ''
            am_1 = "1 AM " if pd.isna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 12)]) and pd.isna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 9)]) and pd.isna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 6)]) and pd.isna(row[('VISA /MASTER/AMEX MACRO Desde el 9/5', 3)]) else ''
            
            #ACS
            acs_6 = "6 ACS " if pd.notna(row[('AMEX GALICIA/PATAG/MACRO/SANTAN/HSBC CUOTA SIMPLE', 6)]) else ''
            acs_3 = "3 ACS " if pd.notna(row[('AMEX GALICIA/PATAG/MACRO/SANTAN/HSBC CUOTA SIMPLE', 3)]) else ''
            acs_1 = "1 ACS " if pd.isna(row[('AMEX GALICIA/PATAG/MACRO/SANTAN/HSBC CUOTA SIMPLE', 6)]) and pd.isna(row[('AMEX GALICIA/PATAG/MACRO/SANTAN/HSBC CUOTA SIMPLE', 3)]) else ''
            
            #SL
            sl_14 = "14 SL " if pd.notna(row[('SOL', 'SOL 14')]) else ''
            sl_12 = "12 SL " if pd.notna(row[('SOL', 'Sol    12')]) else ''
            sl_9 = "9 SL " if pd.notna(row[('SOL', 'Sol     9')]) else ''
            sl_6 = "6 SL " if pd.notna(row[('SOL', 'Sol     6')]) else ''
            sl_3 = "3 SL " if pd.notna(row[('SOL', 'Sol    3')]) else ''
            sl_1 = "1 SL " if pd.isna(row[('SOL', 'SOL 14')]) and pd.isna(row[('SOL', 'Sol    12')]) and pd.isna(row[('SOL', 'Sol     9')]) and pd.isna(row[('SOL', 'Sol     6')]) and pd.isna(row[('SOL', 'Sol    3')]) else ''
            
            #N
            n_18 = "18 N " if pd.notna(row[('NARANJA', 'naran 18')]) else ''
            n_14 = "14 N " if pd.notna(row[('NARANJA', 'naran 14')]) else ''
            n_12 = "12 N " if pd.notna(row[('NARANJA', 'naran 12')]) else ''
            n_9 = "9 N " if pd.notna(row[('NARANJA', 'naranja 9')]) else ''
            n_6 = "6 N " if pd.notna(row[('NARANJA', 'naran 6')]) else ''
            n_1 = "1 N " if pd.isna(row[('NARANJA', 'naran 18')]) and pd.isna(row[('NARANJA', 'naran 14')]) and pd.isna(row[('NARANJA', 'naran 12')]) and pd.isna(row[('NARANJA', 'naranja 9')]) and pd.isna(row[('NARANJA', 'naran 6')]) else ''
            
            #S
            s_12 = "12 S " if pd.notna(row[('SU CREDITO', 'su cre 12')]) else ''
            s_9 = "9 S " if pd.notna(row[('SU CREDITO', 'su cre 9')]) else ''
            s_6 = "6 S " if pd.notna(row[('SU CREDITO', 'su cre 6')]) else ''
            s_5 = "5 S " if pd.notna(row[('SU CREDITO', 'su cre 5')]) else ''
            s_3 = "3 S " if pd.notna(row[('SU CREDITO', 'su cre 3')]) else ''
            s_1 = "1 S " if pd.isna(row[('SU CREDITO', 'su cre 12')]) and pd.isna(row[('SU CREDITO', 'su cre 9')]) is pd.isna(row[('SU CREDITO', 'su cre 6')]) and pd.isna(row[('SU CREDITO', 'su cre 5')]) and pd.isna(row[('SU CREDITO', 'su cre 3')]) else ''

            #JSON RESULTADOS
            resultados.append({
                #SKU
                'SKU': sku,

                #CS
                '1 CS': cs_1,
                '3 CS': cs_3,
                '6 CS': cs_6,

                #VMP
                '18 VMP': vmp_18,
                '15 VMP': vmp_15,
                '12 VMP': vmp_12,
                '9 VMP': vmp_9,
                '1 VMP': vmp_1,

                #VMNP
                "18 VMNP": vmnp_18,
                "15 VMNP": vmnp_15,
                "12 VMNP": vmnp_12,
                "9 VMNP": vmnp_9,
                "1 VMNP": vmnp_1,

                #AM
                "12 AM": am_12,
                "9 AM": am_9,
                "6 AM": am_6,
                "3 AM": am_3,
                "1 AM": am_1,

                #ACS
                "6 ACS": acs_6,
                "3 ACS": acs_3,
                "1 ACS": acs_1,

                #SL
                "14 SL": sl_14,
                "12 SL": sl_12,
                "9 SL": sl_9,
                "6 SL": sl_6,
                "3 SL": sl_3,
                "1 SL": sl_1,

                #N
                "18 N": n_18,
                "14 N": n_14,
                "12 N": n_12,
                "9 N": n_9,
                "6 N": n_6,
                "1 N": n_1,

                #S
                "12 S": s_12,
                "9 S": s_9,
                "6 S": s_6,
                "5 S": s_5,
                "3 S": s_3,
                "1 S": s_1,

                'Condicion_comercial': f'DH {cs_3}{cs_6}{cs_1}{vmp_18}{vmp_15}{vmp_12}{vmp_9}{vmp_1}{vmnp_18}{vmnp_15}{vmnp_12}{vmnp_9}{vmnp_1}{am_12}{am_9}{am_6}{am_3}{am_1}{acs_6}{acs_3}{acs_1}{sl_14}{sl_12}{sl_9}{sl_6}{sl_3}{sl_1}{n_18}{n_14}{n_12}{n_9}{n_6}{n_1}{s_12}{s_9}{s_6}{s_5}{s_3}{s_1}'.strip()

            })

        df_resultado = pd.DataFrame(resultados)
        
        # Generar nombre de archivo único
        filename = f"exportacion_cc_{uuid.uuid4()}.xlsx"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        
        # Guardar el archivo
        df_resultado[['SKU', 'Condicion_comercial']].to_excel(filepath, index=False, sheet_name='Resultados')
        
        flash('Archivo generado correctamente. Puedes descargarlo debajo.', 'success')

        print("UPLOAD_FOLDER:", UPLOAD_FOLDER)
        print("Archivos en carpeta:", os.listdir(UPLOAD_FOLDER))

        
        # Renderizar la plantilla con el botón de descarga
        return render_template('index.html', download_filename=filename)

    except Exception as e:
        flash(f'Error al procesar el archivo: {e}', 'danger')
        return redirect(url_for('home.home'))