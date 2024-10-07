import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import getpass



while True:

    gonderici_email = " " # fill the blank
    gonderici_sifre = " " # fill the blank

    devam = input("Emin misiniz? (Evet/Hayır): ").lower().strip()
    if devam == "evet":
        break
    elif devam != "hayır":
        print("Lütfen sadece 'Evet' veya 'Hayır' şeklinde yanıt verin.")

print("E-posta ve şifre kontrol ediliyor...")

mail_sunucu = "outgoing.itu.edu.tr"
mail_port = 587
kullanici_adi = gonderici_email.split("@")[0]
df = pd.read_excel('epostalar.xlsx')

def mail_gonder(df):
    try:
        server = smtplib.SMTP(mail_sunucu, mail_port)
        server.starttls()
        server.login(kullanici_adi, gonderici_sifre)

        column_names = df.columns.tolist()
        param_vars = {col_name: df[col_name] for col_name in column_names if isinstance(col_name, str) and col_name.startswith("Param_")}
        cc = [" "] # fill the blank

        cckontrol = input("CC gönderilsin mi? (Evet/Hayır): ").lower().strip()
        if cckontrol == "evet":
            for index, row in df.iterrows():
                email_address = row["E-posta"]
                mail_from = row["Gönderen"]
                mail_subject = row["Konu"]
                text = row["Metin"]

                if not pd.isna(email_address):
                    msg = MIMEMultipart()
                    msg['From'] = f"{mail_from} <" + gonderici_email + ">"
                    msg['To'] = email_address
                    msg['Subject'] = mail_subject
                    msg['Cc'] = ", ".join(cc)
                    message_text = text
                    for param_name, param_value in param_vars.items():
                        message_text = message_text.replace(f"{param_name}", str(param_value[index]))
                
                    msg.attach(MIMEText(message_text, 'plain'))
                    server.sendmail(gonderici_email, [email_address] + cc, msg.as_string())
                    print(f"E-posta {email_address} adresine başarıyla gönderildi.")
                else:
                    print("E-posta adresi bulunamadığı için gönderilemedi.")
        
        elif cckontrol == "hayır":
            for index, row in df.iterrows():
                email_address = row["E-posta"]
                mail_from = row["Gönderen"]
                mail_subject = row["Konu"]
                text = row["Metin"]

                if not pd.isna(email_address):
                    msg = MIMEMultipart()
                    msg['From'] = f"{mail_from} <" + gonderici_email + ">"
                    msg['To'] = email_address
                    msg['Subject'] = mail_subject
                    message_text = text
                    for param_name, param_value in param_vars.items():
                        message_text = message_text.replace(f"{param_name}", str(param_value[index]))
                
                    msg.attach(MIMEText(message_text, 'plain'))
                    server.sendmail(gonderici_email, email_address, msg.as_string())
                    print(f"E-posta {email_address} adresine başarıyla gönderildi.")
                else:
                    print("E-posta adresi bulunamadığı için gönderilemedi.")
                       
        server.quit()
    except Exception as e:
        print("E-posta gönderme hatası:", e)

if __name__ == "__main__":
    mail_gonder(df)