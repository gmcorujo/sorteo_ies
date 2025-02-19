from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import random
from fpdf import FPDF
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

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
                               pre_inscriptos=pre_inscriptos,  # Asegúrate de pasar esta variable
                               reservas_count=reservas)       # Asegúrate de pasar esta variable

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
    preinscritos = request.form.getlist('preinscritos[]')
    reservas = request.form.getlist('reservas[]')

    # Convertir los datos a un DataFrame
    df_preinscritos = pd.DataFrame(eval(preinscritos[0]))
    df_reservas = pd.DataFrame(eval(reservas[0]))

    # Crear un archivo Excel con múltiples hojas
    with pd.ExcelWriter('resultados.xlsx', engine='openpyxl') as writer:
        df_preinscritos.to_excel(writer, sheet_name='Preinscritos', index=False)
        df_reservas.to_excel(writer, sheet_name='Reservas', index=False)

    return send_file('resultados.xlsx', as_attachment=True)

# Ruta para descargar el archivo PDF
@app.route('/descargar_pdf', methods=['POST'])
def descargar_pdf():
    preinscritos = request.form.getlist('preinscritos[]')
    reservas = request.form.getlist('reservas[]')

    # Crear un PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)

    # Añadir preinscritos al PDF
    pdf.cell(200, 10, txt="Preinscritos", ln=True, align='C')
    for participante in eval(preinscritos[0]):
        pdf.cell(200, 10, txt=f"{participante['Numero']} - {participante['Nombre']} {participante['Apellido']}", ln=True)

    # Añadir reservas al PDF
    pdf.cell(200, 10, txt="Reservas", ln=True, align='C')
    for participante in eval(reservas[0]):
        pdf.cell(200, 10, txt=f"{participante['Numero']} - {participante['Nombre']} {participante['Apellido']}", ln=True)

    # Guardar el PDF
    pdf.output("resultados.pdf")
    return send_file('resultados.pdf', as_attachment=True)


# Ejecutar la aplicación
if __name__ == '__main__':
    app.run(debug=True)