import os
import time
import datetime
import logging
# Conexión SQL
import pyodbc
# Envío de emails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# Desencriptación contraseña
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend

# Conexión a la base de datos
servidor = 'DESKTOP-97GRPND\\SQLEXPRESS'
base_de_datos = 'ADACSCSD'
usuario = 'EnviadorMails'
contraseña = 'test123'

# Configuración de logging
logging.basicConfig(filename='error.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

try:    
    conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={servidor};DATABASE={base_de_datos};UID={usuario};PWD={contraseña}')
    cursor = conn.cursor()
except Exception as e:
    logging.error("Error al conectarse a la base de datos: %s", e)
    raise

try:   
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=1303")
    registro = cursor.fetchone()
    if registro is not None:      
        servidor_smtp = registro.CADENA
        if not servidor_smtp:
            raise ValueError("No se cargó el servidor SMTP en el parámetro 1303 (Servidor SMTP).")
    else:
        raise ValueError("No se encontró el parámetro 1303 (Servidor SMTP).")
except ValueError as ve:
    logging.error("Error: %s", ve)
    raise
except Exception as e:   
    logging.error("Error al buscar el parámetro 1303 (Servidor SMTP) en PARAMETROS: %s", ve)
    raise

try:   
    cursor.execute("SELECT NUMERO FROM PARAMETROS WHERE CODIGO=1304")
    registro = cursor.fetchone()
    if registro is not None:
        puerto_smtp = int(registro.NUMERO)
        if not puerto_smtp:
            raise ValueError("No hay puerto SMTP cargado en el parámetro 1304.")           
    else:
        raise ValueError("No se encontró el parámetro 1304 (Puerto SMTP).")
except ValueError as ve:
    logging.error("Error: %s", ve)
    raise
except Exception as e:
    logging.error("Error al buscar el parámetro 1304 (Puerto SMTP) en PARAMETROS: %s", ve)
    raise

try:
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=1316")
    registro = cursor.fetchone()
    if registro is not None:
        email_remitente = registro.CADENA
        if not email_remitente:
            raise ValueError("No se cargó el mail en el parámetro 1316 (Email del remitente).")
    else:
        raise ValueError("No se encontró el parámetro 1316 (Usuario del remitente).")
except ValueError as ve:
    logging.error("Error: %s", ve)
    raise
except Exception as e:
    logging.error("Error al buscar el parámetro 1316 (Usuario del remitente) en PARAMETROS: %s", ve)
    raise

try:
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=9017")
    registro = cursor.fetchone()
    if registro is not None:
        contraseña_remitente = registro.CADENA
        if not contraseña_remitente:
            raise ValueError("No se cargó la contraseña en el parámetro 9017 (Contraseña del remitente).")
    else:
        raise ValueError("No se encontró el parámetro 9017 (Contraseña del remitente).")
except ValueError as ve:
    logging.error("Error: %s", ve)
    raise
except Exception as e:
    logging.error("Error al buscar el parámetro 9017 (Contraseña del remitente) en PARAMETROS: %s", e)
    raise

def enviar_email(destinatario, asunto, cuerpo, archivo_adjunto=None):
    email = MIMEMultipart()
    email['Subject'] = asunto
    email['From'] = email_remitente
    email['To'] = destinatario

    email.attach(MIMEText(cuerpo))

    if archivo_adjunto:
        with open(archivo_adjunto, 'rb') as arch_adj:
            adjunto = MIMEBase('application', 'octet-stream')
            adjunto.set_payload(arch_adj.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(archivo_adjunto)}')
        email.attach(adjunto)

    try:
        with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
            servidor.starttls()
            servidor.login(email_remitente, contraseña_remitente)
            servidor.sendmail(email_remitente, destinatario, email.as_string())
        
        cursor.execute(f"UPDATE EMAILSLOG SET EST=1 WHERE ANR={registro.ANR}")
        conn.commit()
    except Exception as e:
        print("Error al enviar el mail:", e)

def main():
    while True:
        try:
            cursor.execute(f"SELECT ANR, DST, ASU, CUE, ADJ FROM EMAILSLOG WHERE EST=0")
            registros = cursor.fetchall()

            for registro in registros:

                envio_fecha_hora = datetime.datetime.now()

                fecha_inicial = datetime.date(1800, 12, 28) # En Clarion el 28/12/1800 es la fecha 1
                fecha_envio_numerico = (envio_fecha_hora.date() - fecha_inicial)
                
                # En Clarion 1 minuto equivale a 6001.
                # Para obtener la hora hay que multiplicar 6001 x 60 x Cantidad Horas 
                # Para obtener los minutos hay que multiplicar 6001 x Cantidad Minutos
                # Para obtener los segundos hay que multiplicar 100 x Cantidad Segundos
                tiempo_envio = ((envio_fecha_hora.hour) * 60 * 6001) + ((envio_fecha_hora.minute) * 6001) + (envio_fecha_hora.second * 100)

                try:
                    cursor.execute(f"UPDATE EMAILSLOG SET EST=99, ENVFEC=?, ENVHOR=? WHERE ANR={registro.ANR}", fecha_envio_numerico, tiempo_envio)
                except Exception as e:
                    print("Error al intentar actualizar el estado del registro al estado 99 (en proceso) y cambiar hora y fecha")

                destinatario = registro.DST
                asunto = registro.ASU
                cuerpo = registro.CUE
                archivo_adjunto = registro.ADJ

                enviar_email(destinatario, asunto, cuerpo, archivo_adjunto)

            # Debería hacer el close acá o al salir del While?
            #conn.close()
        except Exception as e:
            print("2 - Error al enviar el mail:", e) 

            codigo_error = e.args[0] if e.args else None
            if codigo_error == 2:
                mensaje_error = "Verifique la existencia del archivo adjunto."
            else:
                mensaje_error = e.args[1]

            try:
                cursor.execute(f"UPDATE EMAILSLOG SET EST=?, ERRCOD=?, ERRDES=? WHERE ANR={registro.ANR}", (12, codigo_error, mensaje_error))     
                conn.commit()
            except Exception as e:
                print("3 - Error al modificar el registro en EMAILSLOG:", e)    
                 
        break
        time.sleep(15)

    # Cerrar conexión SQL ¿Deberia cerrarla acá? ¿O después de terminar el loop y antes del SELECT hago una nueva conexión?        
    conn.close()

if __name__ == "__main__":
    main()