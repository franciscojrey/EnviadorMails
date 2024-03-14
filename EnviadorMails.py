import os
import time
import datetime
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

try:
    conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={servidor};DATABASE={base_de_datos};UID={usuario};PWD={contraseña}')
    cursor = conn.cursor()
except Exception as e:
    print("Error al conectarse a la base de datos:", e)

try:   
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=1303")
    registro = cursor.fetchone()
    if registro:       
        servidor_smtp = registro.CADENA
    else:
        print("No se encontró el parámetro 1303 (Servidor SMTP).")
except Exception as e:
    print("Error al buscar el parámetro 1303 (Servidor SMTP) en PARAMETROS:", e)

try:   
    cursor.execute("SELECT NUMERO FROM PARAMETROS WHERE CODIGO=1304")
    registro = cursor.fetchone()
    if registro:       
        puerto_smtp = int(registro.NUMERO)
    else:
        print("No se encontró el parámetro 1304 (Puerto SMTP).")
except Exception as e:
    print("Error al buscar el parámetro 1304 (Puerto SMTP) en PARAMETROS:", e)

try:   
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=9017")
    registro = cursor.fetchone()
    if registro:       
        contraseña_remitente = registro.CADENA
    else:
        print("No se encontró el parámetro 9017 (Contraseña del remitente).")
except Exception as e:
    print("Error al buscar el parámetro 9017 (Contraseña del remitente) en PARAMETROS:", e)

email_remitente = 'franciscorey98@gmail.com'

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
    except Exception as e:
        print("Error al enviar el mail:", e)

def main():
    while True:
        try:
            cursor.execute(f"SELECT ANR, DST, ASU, CUE, ADJ FROM EMAILSLOG WHERE EST=0")
            registros = cursor.fetchall()

            for registro in registros:
                destinatario = registro.DST
                asunto = registro.ASU
                cuerpo = registro.CUE
                archivo_adjunto = registro.ADJ
                #sent_datetime = datetime.datetime.now()       
                #sent_date = sent_time.date()
                #sent_time = sent_time.time()
                enviar_email(destinatario, asunto, cuerpo, archivo_adjunto)

                cursor.execute(f"UPDATE EMAILSLOG SET EST=1 WHERE ANR={registro.ANR}")
                conn.commit()

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