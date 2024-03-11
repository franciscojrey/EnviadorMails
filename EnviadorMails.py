import time
# Conexión SQL
import pyodbc
# Envío de emails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# Desencriptación
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend

# Conexión a la base de datos
server = 'DESKTOP-97GRPND\SQLEXPRESS'
database = 'ADACSCSD'
username = 'EnviadorMails'
password = 'test123'
try:
    conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')
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

# Esto no está funcionando
def decrypt_password(encoded_str):
    decrypted_str = ""
    for pos, char in enumerate(encoded_str):
        xor_key = 255 - ((pos + 1) * 2) % 32
        decrypted_char = chr(ord(char) ^ xor_key)
        decrypted_str += decrypted_char
    return decrypted_str

try:   
    cursor.execute("SELECT CADCRYPT FROM PARAMETROS WHERE CODIGO=1306")
    row = cursor.fetchone()
    if row:       
        encrypted_password = row.CADCRYPT
        decrypted_password = decrypt_password(encrypted_password)
        # No es la contraseña correcta parece
        sender_password = decrypt_password
    else:
        print("Contraseña no encontrada en el parámetro 1306.")
except Exception as e:
    print("Error al buscar el parámetro 1306 (contraseña) en PARAMETROS:", e)

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
                send_email(email_destinatario, email_subject, email_body, email_attachment)

                cursor.execute(f"UPDATE EMAILSLOG SET EST=1 WHERE ANR={row.ANR}")
                conn.commit()

            # Debería hacer el close acá o al salir del While?
            #conn.close()
        except Exception as e:

            cursor.execute(f"UPDATE EMAILSLOG SET EST=12 WHERE ANR={row.ANR}")
            conn.commit()

            print("2 - Error:", e)

        break
        time.sleep(15)

    # Cerrar conexión SQL ¿Deberia cerrarla acá? ¿O después de terminar el loop y antes del SELECT hago una nueva conexión?        
    conn.close()

if __name__ == "__main__":
    main()