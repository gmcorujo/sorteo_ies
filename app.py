from flask import Flask, render_template, request, redirect, url_for, send_file

import pandas as pd
import random
import os
import json

from fpdf import FPDF

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    # Limpia espacios en headers y mapea nombres comunes (con acentos / variantes)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        'Número': 'Numero',
        'Numero': 'Numero',
        'Apellido': 'Apellido',
        'Nombre': 'Nombre',
        'Nombres': 'Nombre',

        'Número de Documento': 'Documento',
        'Documento': 'Documento',
        'DNI': 'Documento',

        'Número de Teléfono': 'Telefono',
        'Teléfono': 'Telefono',
        'Telefono': 'Telefono',
        'Celular': 'Telefono',

        'Correo Electrónico': 'Correo',
        'Correo Electrónico ': 'Correo',
        'Email': 'Correo',
        'E-mail': 'Correo',
        'Correo': 'Correo',
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in ['Numero', 'Apellido', 'Nombre', 'Documento', 'Telefono', 'Correo']:
        if col not in df.columns:
            df[col] = ""

    return df.fillna("")


def seleccionar_columnas_export(df: pd.DataFrame) -> pd.DataFrame:
    cols = ['Numero', 'Apellido', 'Nombre', 'Documento', 'Telefono', 'Correo']
    out = df.loc[:, cols].copy()
    return out.rename(columns={
        'Numero': 'Número',
        'Apellido': 'Apellido',
        'Nombre': 'Nombres',
        'Documento': 'Número de Documento',
        'Telefono': 'Teléfono',
        'Correo': 'Correo Electrónico',
    })

# Crear carpeta para subir archivos si no existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Ruta principal
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Obtener los valores del formulario
        try:
            pre_inscriptos = int(request.form.get('pre_inscriptos', 60))
            reservas = int(request.form.get('reservas', 20))
        except ValueError:
            return render_template('index.html', error="Por favor, ingresa números válidos para preinscritos y reservas.")

        # Verificar si se subió un archivo
        if 'archivo' not in request.files:
            return render_template('index.html', error="No se seleccionó ningún archivo.")
        
        archivo = request.files['archivo']
        if archivo.filename == '':
            return render_template('index.html', error="El archivo está vacío.")
        
        # Guardar el archivo en la carpeta de uploads
        ruta_archivo = os.path.join(app.config['UPLOAD_FOLDER'], archivo.filename)
        archivo.save(ruta_archivo)

        # Leer el archivo Excel
        try:
            df = pd.read_excel(ruta_archivo)

            df = normalizar_columnas(df)
            
            # Reemplazar NaN con cadenas vacías
            df = df.fillna("")
        except Exception as e:
            return render_template('index.html', error=f"Error al leer el archivo: {e}")

        # Realizar el sorteo
        preinscritos, reservas_seleccionadas = realizar_sorteo(df, pre_inscriptos, reservas)
        if preinscritos is None or reservas_seleccionadas is None:
            return render_template('index.html', error="No hay suficientes participantes para el sorteo.")

        # Pasar los resultados y los valores de preinscritos/reservas al template
        return render_template('resultado.html', 
                               preinscritos=preinscritos, 
                               reservas=reservas_seleccionadas,
                               pre_inscriptos=pre_inscriptos,
                               reservas_count=reservas)

    return render_template('index.html')

# Función para realizar el sorteo
def realizar_sorteo(df, pre_inscriptos, reservas):
    if len(df) < pre_inscriptos + reservas:
        return None, None

    # Mezclar aleatoriamente los participantes
    participantes = df.sample(frac=1, random_state=random.randint(1, 1000)).reset_index(drop=True)

    # Separar en preinscritos y reservas
    preinscritos = participantes.iloc[:pre_inscriptos]
    reservas_seleccionadas = participantes.iloc[pre_inscriptos:pre_inscriptos + reservas]

    return preinscritos.to_dict(orient='records'), reservas_seleccionadas.to_dict(orient='records')

# Ruta para descargar el archivo Excel
@app.route('/descargar_excel', methods=['POST'])
def descargar_excel():
    preinscritos_json = request.form.get('preinscritos_json', '[]')
    reservas_json = request.form.get('reservas_json', '[]')

    df_preinscritos = pd.DataFrame(json.loads(preinscritos_json))
    df_reservas = pd.DataFrame(json.loads(reservas_json))

    df_preinscritos = seleccionar_columnas_export(normalizar_columnas(df_preinscritos))
    df_reservas = seleccionar_columnas_export(normalizar_columnas(df_reservas))

    with pd.ExcelWriter('resultados.xlsx', engine='openpyxl') as writer:
        df_preinscritos.to_excel(writer, sheet_name='Preinscritos', index=False)
        df_reservas.to_excel(writer, sheet_name='Reservas', index=False)

    return send_file('resultados.xlsx', as_attachment=True)


# Ruta para descargar el archivo PDF
@app.route('/descargar_pdf', methods=['POST'])
def descargar_pdf():
    preinscritos_json = request.form.get('preinscritos_json', '[]')
    reservas_json = request.form.get('reservas_json', '[]')

    df_pre = seleccionar_columnas_export(normalizar_columnas(pd.DataFrame(json.loads(preinscritos_json))))
    df_res = seleccionar_columnas_export(normalizar_columnas(pd.DataFrame(json.loads(reservas_json))))

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10)

    def imprimir_seccion(titulo, df):
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 8, txt=titulo, ln=True, align='C')
        pdf.ln(2)

        pdf.set_font("Arial", style="B", size=9)
        pdf.multi_cell(0, 5, txt=" | ".join(df.columns))
        pdf.set_font("Arial", size=9)

        for _, row in df.iterrows():
            pdf.multi_cell(0, 5, txt=" | ".join(str(row[c]) for c in df.columns))
        pdf.ln(2)

    imprimir_seccion("Preinscritos", df_pre)
    imprimir_seccion("Reservas", df_res)

    pdf.output("resultados.pdf")
    return send_file("resultados.pdf", as_attachment=True)

# Ejecutar la aplicación
if __name__ == '__main__':
    app.run(debug=True)