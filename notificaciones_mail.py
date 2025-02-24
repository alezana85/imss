import imaplib
import email
from telegram import Bot
import time
import asyncio
import keyboard
import sys

# Configuración de Gmail
GMAIL_USER = "hri.notificaciones@gmail.com"
GMAIL_PASSWORD = "waki gpdk nwlo mjgc"
IMAP_SERVER = "imap.gmail.com"

# Configuración de Telegram
TOKEN = "8087764869:AAGkWUQZx_eFbOjpPlTXVW_Woi3gk7qSV18"
CHAT_ID = "6894155711"

# Lista de palabras clave
KEYWORDS = [
    "TERRAM", "ARMONICA", "ASESORES", "ATOMICA", "AVL", "BALA", "BAVARIE",
    "BUSINESS", "CFO", "COLEGIO", "COMER45", "COMERDEMEX", "MAGAVAZ",
    "ALUMBRADO", "OPPENTANO", "CUINBA", "EFRAIN", "WTRACKING", "FANCICOM",
    "GESTION DE PLANES", "GOBUA", "GP MACHINERY", "KIMISA", "HENESI", "HRI GLOBAL",
    "INFRA", "ORBITAL", "FALCOMERX", "INTELLISWITCH", "INVERSION BURSATIL",
    "IP TIGO", "KOCARE", "LA PIEDRA", "LEARNING", "MD ILUMINACION",
    "MAR Y TIERRA", "NEUMATICOS", "PIXEL", "PLASBRI", "JOPLA", "A TU ALCANCE",
    "CALORZA", "SERVICIOS ATOMICA", "MONTEFOR", "TECNOLOGY", "KENA",
    "VM CONSULTORES", "VM CONSULTORIA", "VOLCAN 222", "SEPARACIONES", "SEPARACION"
]

async def send_telegram_message(message):
    bot = Bot(token=TOKEN)
    await bot.send_message(chat_id=CHAT_ID, text=message)

def decode_text(text_bytes):
    encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
    for encoding in encodings:
        try:
            return text_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return "No se pudo decodificar el contenido del correo"

async def check_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(GMAIL_USER, GMAIL_PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, "UNSEEN")
    if status != "OK":
        return

    for num in messages[0].split():
        status, data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            continue

        msg = email.message_from_bytes(data[0][1])
        subject = msg["Subject"] or ""
        body = ""

        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        body = decode_text(payload)
                        break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = decode_text(payload)

        for keyword in KEYWORDS:
            if keyword.lower() in subject.lower() or keyword.lower() in body.lower():
                await send_telegram_message(f"¡Recibiste un correo de {keyword}!\nAsunto: {subject}")
                break

    mail.close()
    mail.logout()

async def main():
    print("Script iniciado. Presiona ESC para detener.")
    try:
        while True:
            if keyboard.is_pressed('esc'):
                print("\nDetención solicitada. Cerrando script...")
                sys.exit(0)
            await check_emails()
            await asyncio.sleep(120)  # Verificar cada 2 minutos
    except KeyboardInterrupt:
        print("\nDetención solicitada. Cerrando script...")
    except SystemExit:
        pass
    except Exception as e:
        print(f"Error inesperado: {e}")
    finally:
        print("Script finalizado.")

if __name__ == "__main__":
    asyncio.run(main())