import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# Database connection settings
server = 'DESKTOP-97GRPND\SQLEXPRESS'
database = 'ADACSCSD'
username = 'EnviadorMails'
password = 'test123'

# Email settings
smtp_server = 'smtp.gmail.com'
smtp_port = 587
sender_email = 'franciscorey98@gmail.com'
sender_password = ''
receiver_email = 'franciscorey98@gmail.com'

# Connect to the database
try:
    conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    cursor = conn.cursor()
except Exception as e:
    print("Error al conectarse a la base de datos:", e)

def send_email(subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email

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
        server.sendmail(sender_email, receiver_email, msg.as_string())

def main():
    while True:
        try:
            # Query for records with "Sent" column equals 0
            cursor.execute(f"SELECT * FROM EMAILSLOG WHERE EST=0")
            rows = cursor.fetchall()

            for row in rows:
              
                # Extract information from other columns
                email_subject = f"{row[6]}"
                email_body = f"{row[7]}"
                email_attachment = f"{row[8]}"

                # Send email
                send_email(email_subject, email_body, email_attachment)

                # Update "EST" column to 1 for the processed record
                cursor.execute(f"UPDATE EMAILSLOG SET EST=1 WHERE ANR={row[0]}")
                conn.commit()

            # Debería hacer el close acá o al salir del While?
            #conn.close()
        except Exception as e:
            print("Error:", e)

        # Esperar 15 seg
        time.sleep(15)

    # Cerrar conexión SQL ¿Deberia cerrarla acá? ¿O después de terminar el loop y antes del SELECT hago una nueva conexión?        
    conn.close()

if __name__ == "__main__":
    main()