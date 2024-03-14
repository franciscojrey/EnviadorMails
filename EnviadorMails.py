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
    row = cursor.fetchone()
    if row:       
        smtp_server = row.CADENA
    else:
        print("SMTP server with parameter number 1303 not found")
except Exception as e:
    print("Error al buscar el parámetro 1303 (SMTP SERVER) en PARAMETROS:", e)

try:   
    cursor.execute("SELECT NUMERO FROM PARAMETROS WHERE CODIGO=1304")
    row = cursor.fetchone()
    if row:       
        smtp_port = int(row.NUMERO)
    else:
        print("El puerto en el parámetro 1304 no fue encontrado.")
except Exception as e:
    print("Error al buscar el parámetro 1304 (puerto) en PARAMETROS:", e)

try:   
    cursor.execute("SELECT CADENA FROM PARAMETROS WHERE CODIGO=9017")
    row = cursor.fetchone()
    if row:       
        sender_password = row.CADENA
    else:
        print("Contraseña no encontrada en el parámetro 9017.")
except Exception as e:
    print("Error al buscar el parámetro 9017 (contraseña) en PARAMETROS:", e)

sender_email = 'franciscorey98@gmail.com'

def send_email(destinatario, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = destinatario

    msg.attach(MIMEText(body))

    if attachment_path:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {attachment_path}')
        msg.attach(part)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, destinatario, msg.as_string())

def main():
    while True:
        try:
            # Query for records with "Sent" column equals 0
            cursor.execute(f"SELECT ANR, DST, ASU, CUE, ADJ FROM EMAILSLOG WHERE EST=0")
            rows = cursor.fetchall()

            for row in rows:
                email_destinatario = row.DST
                email_subject = row.ASU
                email_body = row.CUE
                email_attachment = row.ADJ
                #sent_datetime = datetime.datetime.now()       
                #sent_date = sent_time.date()
                #sent_time = sent_time.time()
                send_email(email_destinatario, email_subject, email_body, email_attachment)

                cursor.execute(f"UPDATE EMAILSLOG SET EST=1 WHERE ANR={row.ANR}")
                conn.commit()

            # Debería hacer el close acá o al salir del While?
            #conn.close()
        except Exception as e:
            print("2 - Error al enviar el mail:", e) 

            error_code = e.args[0] if e.args else None
            if error_code == 2:
                error_message = "Verifique la existencia del archivo adjunto."
            else:
                error_message = e.args[1]

            try:
                cursor.execute(f"UPDATE EMAILSLOG SET EST=?, ERRCOD=?, ERRDES=? WHERE ANR={row.ANR}", (12, error_code, error_message))     
                conn.commit()
            except Exception as e:
                print("3 - Error al modificar el registro en EMAILSLOG:", e)    
                 
        break
        time.sleep(15)

    # Cerrar conexión SQL ¿Deberia cerrarla acá? ¿O después de terminar el loop y antes del SELECT hago una nueva conexión?        
    conn.close()

if __name__ == "__main__":
    main()