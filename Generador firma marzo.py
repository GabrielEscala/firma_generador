import os
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, Text, messagebox
from jinja2 import Template
import logging

# Create a logger
logger = logging.getLogger('signature_generator')
logger.setLevel(logging.INFO)

# Create a file handler
file_handler = logging.FileHandler('signature_generator_convencional.log')
file_handler.setLevel(logging.INFO)

# Create a console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Create a formatter and set it for the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add the handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Número de teléfono fijo
fixed_phone_number = "+1(305)707.6761"

# Plantilla HTML para la firma convencional (con ajustes para Outlook)
signature_template_html = """
<!DOCTYPE html>
<html>
<head>
  <meta content="text/html; charset=UTF-8" http-equiv="content-type">
</head>
<body style="margin: 0; padding: 0; font-family: 'Helvetica Neue', Arial, sans-serif;">
  <!-- Contenedor principal con ancho fijo para alineación perfecta -->
  <table width="650" cellpadding="0" cellspacing="0" border="0" 
  style="width:650px; max-width:650px; margin:0; padding:0; font-family:'Helvetica Neue', Arial, sans-serif; border-collapse:collapse;">

  <tr>
  <td style="padding:0; margin:0;">

        <!-- Bloque para el nombre y título -->
        <p style="margin: 0; padding: 0; font-size: 14pt; font-family: 'Helvetica Neue', sans-serif; color: #545454;">
          {{ name }} | <span style="font-family: HelveticaNeue-Light, sans-serif; color: #545454;">{{ title }}</span>
        </p>

        <!-- Espaciador extra para que Outlook separe nombre/cargo de la línea del correo -->
        <table width="650" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td style="height: 10px; line-height: 10px; font-size: 1px;">&nbsp;</td>
          </tr>
        </table>

        <!-- Bloque para el email, teléfono, nombre, y URL usando tabla para mayor compatibilidad y separación -->
       <table width="650" cellpadding="0" cellspacing="0" border="0" style="width: 650px; table-layout: fixed; border-collapse: collapse; font-family: Arial, sans-serif;">
  <tr>
    <td width="215" style="width: 215px; padding: 0; vertical-align: middle;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="padding-right: 6px;"><img src="https://res.cloudinary.com/dbr7ugqh6/image/upload/v1749049928/Mail_cs1pet.png" width="17" height="17" style="display:block;"></td>
          <td style="font-size: 9pt; color: black;"><a href="mailto:{{ email }}" style="color: black; text-decoration: none;">{{ email }}</a></td>
        </tr>
      </table>
    </td>

    <td width="150" style="width: 150px; padding: 0; vertical-align: middle;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="padding-right: 6px;"><img src="https://res.cloudinary.com/dbr7ugqh6/image/upload/v1749049928/Cellphone_uoctej.png" width="17" height="17" style="display:block;"></td>
          <td style="font-size: 9pt; color: black;"><a href="tel:{{ phone }}" style="color: black; text-decoration: none;">{{ phone }}</a></td>
        </tr>
      </table>
    </td>

    <td width="135" style="width: 135px; padding: 0; vertical-align: middle;">
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="padding-right: 6px;"><img src="https://res.cloudinary.com/dbr7ugqh6/image/upload/v1749049928/Teams_mtnr3x.png" width="17" height="17" style="display:block;"></td>
          <td style="font-size: 9pt; color: black;">{{ name }}</td>
        </tr>
      </table>
    </td>

    <td width="150" style="width: 150px; padding: 0; vertical-align: middle;">
      <table cellpadding="0" cellspacing="0" border="0" align="right">
        <tr>
          <td style="padding-right: 6px;"><img src="https://res.cloudinary.com/dbr7ugqh6/image/upload/v1749049928/web_t4odpo.png" width="17" height="17" style="display:block;"></td>
          <td style="font-size: 9pt; color: black;"><a href="http://www.escalabeds.com" style="color: black; text-decoration: none;">www.escalabeds.com</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><td colspan="4" style="height: 15px; font-size: 1px;">&nbsp;</td></tr>
</table>

        <!-- Resto del contenido, como direcciones y avisos legales -->
        <table width="650" cellpadding="0" cellspacing="0" border="0" style="width: 650px; min-width: 650px; max-width: 650px; border-top: 2pt solid #004f9e; padding-top: 10px; margin: 10px 0 0 0; font-family: Arial, sans-serif; font-size: 8pt; color: #000000; table-layout: fixed; border-collapse: collapse;">
  <tr>
    <td width="110" style="width: 110px; padding: 0; vertical-align: top;">
      <div style="width: 105px; overflow: hidden;">
        <strong style="font-size: 9pt;">Miami, USA</strong><br>
        1000 Brickell Av, 1015<br>
        Miami, FL 33131, US
      </div>
    </td>
    <td width="95" style="width: 95px; padding: 0; vertical-align: top;">
      <div style="width: 90px; overflow: hidden;">
        <strong style="font-size: 9pt;">Caracas, VZLA</strong><br>
        Calle San Felipe,<br>
        Torre Coinasa. PH
      </div>
    </td>
    <td width="115" style="width: 115px; padding: 0; vertical-align: top;">
      <div style="width: 110px; overflow: hidden;">
        <strong style="font-size: 9pt;">Palma de Mallorca</strong><br>
        Gremi Fusters 33, 315<br>
        07009, ESP
      </div>
    </td>
    <td width="105" style="width: 105px; padding: 0; vertical-align: top;">
      <div style="width: 100px; overflow: hidden;">
        <strong style="font-size: 9pt;">Bogotá, COL</strong><br>
        Carrera 15 #93A-84<br>
        8th Floor, Of. 802
      </div>
    </td>
    <td width="125" style="width: 125px; padding: 0; vertical-align: top;">
      <div style="width: 120px; overflow: hidden;">
        <strong style="font-size: 9pt;">CD México, MEX</strong><br>
        Paseo de la Reforma 509<br>
        06500. 16th Floor
      </div>
    </td>
    <td width="100" style="width: 100px; padding: 0; vertical-align: top;">
      <div style="width: 100px; overflow: hidden;">
        <strong style="font-size: 9pt;">London, UK</strong><br>
        54 Portland Place,<br>
        London W1B 1DY
      </div>
    </td>
  </tr>
</table>

        <!-- Espaciador entre ubicaciones y banner (reducido) -->
        <table width="650" cellpadding="0" cellspacing="0" border="0" style="width: 650px; max-width: 650px; margin: 0; padding: 0;">
          <tr>
            <td style="height: 6px; line-height: 6px; font-size: 1px;">&nbsp;</td>
          </tr>
        </table>

        <!-- Imagen / banner -->
        <table width="650" cellpadding="0" cellspacing="0" border="0" style="width: 650px; max-width: 650px; margin: 0; padding: 0; border-collapse: collapse;">
  <tr>
    <td width="650" style="width: 650px; padding: 0; margin: 0; line-height: 1px; font-size: 1px; mso-line-height-rule: exactly; vertical-align: top;">
      <img src="https://nyc.cloud.appwrite.io/v1/storage/buckets/6997191c000394842ec8/files/699e046b0039c6045eeb/view?project=699718eb002f63fcd35e&token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ0b2tlbklkIjoiNjk5ZTA0ODQ1MmZkMGQ0ZGMxODciLCJyZXNvdXJjZUlkIjoiNjk5NzE5MWMwMDAzOTQ4NDJlYzg6Njk5ZTA0NmIwMDM5YzYwNDVlZWIiLCJyZXNvdXJjZVR5cGUiOiJmaWxlcyIsInJlc291cmNlSW50ZXJuYWxJZCI6Ijk4Nzk1OjYiLCJpYXQiOjE3NzE5NjM1MjQsImV4cCI6MTc3NDQ4MzIwMH0.Db8hhcPKabp5FkX6SQvisMpnCxOc04XwkwkCRAX9NgM"
           alt="ESCALABEDS Banner"
           width="650"
           border="0"
           style="display: block; width: 650px; max-width: 650px; height: auto; border: 0; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic;">
    </td>
  </tr>
</table>

        <!-- Avisos legales en párrafos para evitar estiramiento de la última línea en Outlook -->
        <table width="650" cellpadding="0" cellspacing="0" border="0" style="width: 650px; max-width: 650px; border-top: 2pt solid #004f9e; margin: 10px 0 0 0; table-layout: fixed; border-collapse: collapse;">
  <tr>
    <td style="padding-top: 10px; font-family: Arial, sans-serif; font-size: 5pt; color: #000000; line-height: 8pt; text-align: justify; mso-line-height-rule: exactly;">
      
      <div style="width: 650px; margin-bottom: 8px; word-break: break-word;">
        <strong>AVISO LEGAL:</strong> Esta información es privada y confidencial. Se encuentra únicamente dirigida a su(s) destinatario(s). Si usted no es el destinatario original de este mensaje o los documentos adjuntos no tienen relación con usted y por este medio pudo acceder a dicha información, por favor notifique al remitente y elimine el mensaje con su archivo adjunto (si lo hubiere). Esta comunicación es sólo para propósitos informativos y no debe ser considerada como propuesta, aceptación ni como una declaración de voluntad oficial de ESCALABEDS, LLC a menos que ello se señale en forma expresa. Su transmisión no garantiza que el correo electrónico sea seguro o libre de error. Por consiguiente, no manifestamos que esta información sea completa o precisa. Toda información está sujeta a alterarse sin previo aviso.
      </div>

      <div style="width: 650px; word-break: break-word;">
        <strong>LEGAL NOTICE:</strong> The information contained in the present e-mail is private and confidential. It is only addressed to the email recipient(s). If you are not the original recipient of this message or the enclosed documents are not related to you, please notify the sender and delete the message with the enclosed documents (if any). This communication is only for informative purposes and shall not be considered as a proposal, acceptance nor an official statement of intent of ESCALABEDS, LLC, unless expressly indicated. Its transmission does not guarantee that the email is safe or free of error. Therefore, we do not state that this information is complete or accurate. All information is subject to change without notice.
      </div>

    </td>
  </tr>
</table>
      </td>
    </tr>
  </table>
</body>
</html>
"""

def generate_signature(template, data):
    try:
        t = Template(template)
        return t.render(data)
    except Exception as e:
        logger.error(f"Error generating signature: {str(e)}")
        return None

def read_employee_data(file_path):
    try:
        df = pd.read_excel(file_path)
        # Verificar que las columnas necesarias están presentes
        required_columns = ['Nombre para mostrar', 'Título', 'Nombre principal de usuario']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("El archivo Excel no contiene todas las columnas necesarias.")
        logger.info(f"Employee data read from {file_path}")
        return df
    except Exception as e:
        logger.error(f"Error reading employee data: {str(e)}")
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel: {str(e)}")
        return None

def export_signatures(file_path):
    df = read_employee_data(file_path)
    if df is None:
        return

    output_folder = filedialog.askdirectory(title="Seleccionar carpeta de exportación")
    if not output_folder:
        return

    for index, row in df.iterrows():
        data = {
            'name': row['Nombre para mostrar'],
            'title': row['Título'],
            'email': row['Nombre principal de usuario'],
            'phone': fixed_phone_number
        }

        employee_folder = os.path.join(output_folder, row['Nombre para mostrar'].strip())
        os.makedirs(employee_folder, exist_ok=True)

        signature_html = generate_signature(signature_template_html, data)

        # Guardar archivo con extensión .htm
        file_name = f"{row['Nombre para mostrar'].strip()}_Marzo_Onato.htm"
        output_file_path = os.path.join(employee_folder, file_name)

        try:
            with open(output_file_path, "w", encoding='utf-8') as file:
                file.write(signature_html)
            logger.info(f"Signature exported to {output_file_path}")
        except Exception as e:
            logger.error(f"Error exporting signature: {str(e)}")

    logger.info("Signatures exported successfully")
    messagebox.showinfo("Éxito", "Firmas exportadas exitosamente.")

def preview_first_employee(df):
    for index, row in df.iterrows():
        data = {
            'name': row['Nombre para mostrar'],
            'title': row['Título'],
            'email': row['Nombre principal de usuario'],
            'phone': fixed_phone_number
        }
        preview_signature(data)
        break  # Previsualizar solo el primer empleado

def preview_signature(data):
    signature_html = generate_signature(signature_template_html, data)
    preview_text.delete('1.0', 'end')
    preview_text.insert('1.0', signature_html)

def open_file_dialog():
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
    )
    if file_path:
        df = read_employee_data(file_path)
        if df is not None:
            preview_first_employee(df)
            export_signatures(file_path)

# Interfaz gráfica
root = Tk()
root.title("Generador de Firmas Convencionales - ESCALABEDS")

label = Label(root, text="Generador de Firmas Convencionales", font=("Helvetica", 16))
label.pack(pady=20)

button = Button(root, text="Seleccionar archivo Excel para generar firmas", command=open_file_dialog, width=40)
button.pack(pady=10)

preview_label = Label(root, text="Previsualización de Firma", font=("Helvetica", 12))
preview_label.pack(pady=10)

preview_text = Text(root, wrap='word', width=100, height=30)
preview_text.pack(pady=10)

root.mainloop()
