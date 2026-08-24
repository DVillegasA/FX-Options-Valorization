import os
import re
import argparse
import win32com.client as client
import daily_process
from dotenv import load_dotenv

load_dotenv()

outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
attachment_files = []
sender_email = os.getenv("SENDER_EMAIL")
subject_email = os.getenv("SUBJECT_EMAIL")

parser = argparse.ArgumentParser(description="Script para ejecución automática del proceso de Valorización de opciones FX de Nevasa.")
parser.add_argument("--responder", help="Determina si el script debe responder el correo con las carteras automáticamente después de su procesamiento. Por defecto no se responde.", action='store_true')

args = parser.parse_args()

def reply_mail(original_mail, message, attachments: list):
    reply = original_mail.ReplyAll()
    reply.Body = message

    for att in attachments:
        reply.Attachments.Add(att)

    reply.Send()

for i in range(100):
    msg = messages[i]
    if subject_email in msg.Subject and msg.SenderEmailAddress == sender_email and msg.Attachments.Count > 0:
        re_process_date = re.search(r" (\d{4})_(\d{2})_(\d{2})", msg.Subject)
        process_date = f"{re_process_date.group(1)}{re_process_date.group(2)}{re_process_date.group(3)}"
        input_path = os.path.join(os.getcwd(), "data", process_date)
        output_path = os.path.join(os.getcwd(), "salidas")
        
        if not os.path.exists(input_path):
            os.makedirs(input_path)

        if not os.path.exists(output_path):
            os.makedirs(output_path)

        for attachment in msg.Attachments:
            if ".xlsx" not in attachment.FileName:
                continue

            file_path = os.path.join(input_path, attachment.FileName)
            attachment.SaveAsFile(file_path)
            attachment_files.append(attachment.FileName)

        attachment_files.sort()

        if len(attachment_files) > 2:
            print("Error: se descargaron mas archivos de los esperados, favor verifique: ")
            for filename in attachment_files:
                print(filename)
            break

        grid_path = os.path.join(input_path, attachment_files[0])
        portfolio_path = os.path.join(input_path, attachment_files[1])
        
        res = daily_process.procesar(grid_path, portfolio_path, None, output_path)
        print(res["resumen"])
        print("Archivos generados:")
        for a in res["archivos"]:
            print("  ", a)

        reply_message = res["resumen"]
        reply_attachments = [att for att in res["archivos"] if ".xlsx" in att]

        if args.responder:
            reply_mail(msg, reply_message, reply_attachments)
            print("Correo enviado")

        break
