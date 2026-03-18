"""
email_sender.py
---------------
Envía correos personalizados desde Gmail usando datos de un CSV.
Rastrea el estado de cada envío directamente en el CSV.

Uso:
    python email_sender.py

Requisitos:
    pip install pandas
"""

import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import sys

# ==============================================================================
# CONFIGURACIÓN — Edita estos valores antes de ejecutar
# ==============================================================================

GMAIL_USER        = "juancarmona721@gmail.com"       # Tu dirección Gmail
GMAIL_APP_PASSWORD = "ngcw"       # App Password de Google (16 chars)
SENDER_NAME       = "Gracar Solutions"              # Nombre visible del remitente
CSV_FILE          = "psicologosRonda1.csv"               # Ruta al archivo CSV
MAX_EMAILS        = 1                            # Máximo de correos por ejecución
DRY_RUN           = False                        # True = simula sin enviar

# Opciones nuevas basadas en tu correo:
SUBJECT           = "Optimiza tu agenda y reduce inasistencias" # [Asunto - Cambiar si lo deseas]
YOUR_NAME         = "Juan Esteban Carmona"
COMPANY_NAME      = "Gracar Solutions"
WHATSAPP_NUMBER   = "30608"

# ==============================================================================
# PLANTILLA HTML DEL CORREO
# ==============================================================================

HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"></head>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.6;">

<p>Hola {first_name},</p>

<p>Cada cita que no se confirma a tiempo es dinero que se va. Cada mensaje que queda sin responder un domingo es un paciente que busca otro psicólogo.</p>

<p>La pregunta no es si necesita optimizar su agenda. La pregunta es cuánto le está costando no haberlo hecho aún.</p>

<p>Hemos desarrollado un sistema que gestiona toda su agenda por WhatsApp sin que usted o su equipo muevan un dedo:</p>

<p>
→ Pacientes que agendan, confirman y reprograman solos, a cualquier hora<br>
→ Recordatorios automáticos que reducen las inasistencias<br>
→ Su agenda siempre actualizada, sin un solo mensaje manual
</p>

<p>Cuesta una fracción de lo que vale contratar a alguien para hacer lo mismo — y trabaja los 365 días del año sin excepciones.</p>

<p>Si le genera curiosidad, hágame saber y coordinamos una llamada de 20 minutos para que lo vea funcionar. Sin compromiso.</p>

<p>--<br>
<strong>{your_name}</strong><br>
{company_name}<br>
WhatsApp: {whatsapp_number}
</p>

</body>
</html>
"""

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

def extract_first_name(full_name: str) -> str:
    """Extrae el primer nombre de un campo Name como 'Amy W.' → 'Amy'."""
    if not isinstance(full_name, str) or not full_name.strip():
        return "there"  # Saludo genérico si no hay nombre
    return full_name.strip().split()[0]


def load_csv(filepath: str) -> pd.DataFrame:
    """Carga el CSV y agrega columnas de rastreo si no existen."""
    df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str)
    df = df.fillna("")  # Reemplaza NaN por cadena vacía

    # Crear columnas de rastreo si no existen
    for col in ["Status", "Sent_Date", "Error_Log"]:
        if col not in df.columns:
            df[col] = ""

    return df


def save_csv(df: pd.DataFrame, filepath: str) -> None:
    """Guarda el DataFrame de vuelta al CSV preservando encoding."""
    df.to_csv(filepath, index=False, encoding="utf-8-sig")


def build_message(to_email: str, to_name: str, first_name: str) -> MIMEMultipart:
    """Construye el objeto MIMEMultipart con cabeceras y cuerpo HTML."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = SUBJECT
    msg["From"]    = f"{SENDER_NAME} <{GMAIL_USER}>"
    msg["To"]      = to_email

    html_body = HTML_TEMPLATE.format(
        first_name=first_name,
        your_name=YOUR_NAME,
        company_name=COMPANY_NAME,
        whatsapp_number=WHATSAPP_NUMBER
    )
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg


def send_email(smtp_conn: smtplib.SMTP, msg: MIMEMultipart, to_email: str) -> None:
    """Envía el mensaje usando la conexión SMTP activa."""
    smtp_conn.sendmail(GMAIL_USER, to_email, msg.as_bytes())


# ==============================================================================
# LÓGICA PRINCIPAL
# ==============================================================================

def main():
    # ── Cargar CSV ──────────────────────────────────────────────────────────
    try:
        df = load_csv(CSV_FILE)
    except FileNotFoundError:
        print(f"[ERROR] No se encontró el archivo: {CSV_FILE}")
        sys.exit(1)

    total_contacts = len(df)

    # ── Identificar filas pendientes (no enviadas) ───────────────────────────
    already_sent_mask = df["Status"].str.strip().str.lower() == "sent"
    pending_mask      = ~already_sent_mask
    pending_df        = df[pending_mask]

    already_sent_count = already_sent_mask.sum()
    pending_count      = pending_mask.sum()

    print(f"\n{'='*55}")
    print(f"  EMAIL SENDER — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*55}")
    print(f"  Archivo CSV  : {CSV_FILE}")
    print(f"  Total filas  : {total_contacts}")
    print(f"  Ya enviados  : {already_sent_count}")
    print(f"  Pendientes   : {pending_count}")
    print(f"  Límite sesión: {MAX_EMAILS}")
    print(f"  Modo DRY RUN : {DRY_RUN}")
    print(f"{'='*55}\n")

    if pending_count == 0:
        print("[INFO] No hay contactos pendientes. Nada que enviar.")
        return

    # ── Conectar a Gmail SMTP (solo si no es DRY_RUN) ───────────────────────
    smtp_conn = None
    if not DRY_RUN:
        try:
            print("[SMTP] Conectando a smtp.gmail.com:587 …")
            smtp_conn = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
            smtp_conn.ehlo()
            smtp_conn.starttls()
            smtp_conn.ehlo()
            smtp_conn.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            print("[SMTP] Conexión establecida.\n")
        except smtplib.SMTPAuthenticationError:
            print("[ERROR] Fallo de autenticación. Verifica GMAIL_USER y GMAIL_APP_PASSWORD.")
            sys.exit(1)
        except Exception as e:
            print(f"[ERROR] No se pudo conectar al servidor SMTP: {e}")
            sys.exit(1)

    # ── Contadores de sesión ─────────────────────────────────────────────────
    sent_now   = 0
    failed_now = 0

    # ── Iterar sobre filas pendientes ────────────────────────────────────────
    for idx, row in pending_df.iterrows():
        if sent_now >= MAX_EMAILS:
            break

        email_address = str(row.get("Email", "")).strip()
        full_name     = str(row.get("Name", "")).strip()
        sl_no         = str(row.get("SL No", idx)).strip()
        entity        = str(row.get("Pension fund/entity", "")).strip()
        first_name    = extract_first_name(full_name)

        print(f"[{sl_no}] {full_name} <{email_address}> — {entity}")

        # ── Validar email ────────────────────────────────────────────────────
        if not email_address:
            df.at[idx, "Status"]    = "Failed"
            df.at[idx, "Sent_Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "Error_Log"] = "No email address"
            save_csv(df, CSV_FILE)
            print(f"     → [FAILED] Sin dirección de correo.\n")
            failed_now += 1
            continue

        # ── Construir mensaje ────────────────────────────────────────────────
        msg = build_message(email_address, full_name, first_name)

        # ── Enviar (o simular) ───────────────────────────────────────────────
        try:
            if DRY_RUN:
                print(f"     → [DRY RUN] Correo simulado a {email_address} (Hi {first_name})")
            else:
                send_email(smtp_conn, msg, email_address)
                print(f"     → [SENT] Correo enviado a {email_address} (Hi {first_name})")

            df.at[idx, "Status"]    = "Sent"
            df.at[idx, "Sent_Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "Error_Log"] = ""
            sent_now += 1

        except Exception as e:
            error_msg = str(e)
            df.at[idx, "Status"]    = "Failed"
            df.at[idx, "Sent_Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "Error_Log"] = error_msg
            failed_now += 1
            print(f"     → [FAILED] {error_msg}")

        # ── Guardar CSV inmediatamente tras cada intento ─────────────────────
        save_csv(df, CSV_FILE)
        print()

    # ── Cerrar SMTP ──────────────────────────────────────────────────────────
    if smtp_conn:
        try:
            smtp_conn.quit()
        except Exception:
            pass

    # ── Resumen final ────────────────────────────────────────────────────────
    remaining_pending = (df["Status"].str.strip().str.lower() != "sent").sum()
    # Descontar fallidos del pendiente real
    failed_total = (df["Status"].str.strip().str.lower() == "failed").sum()
    still_pending = (df["Status"].str.strip() == "").sum()

    print(f"{'='*55}")
    print(f"  RESUMEN FINAL")
    print(f"{'='*55}")
    print(f"  Total contactos   : {total_contacts}")
    print(f"  Ya enviados antes : {already_sent_count}")
    print(f"  Enviados ahora    : {sent_now}")
    print(f"  Fallidos (total)  : {failed_total}")
    print(f"  Pendientes        : {still_pending}")
    print(f"{'='*55}\n")

    if still_pending > 0:
        print(f"[INFO] Ejecuta el script de nuevo para continuar con los {still_pending} pendientes.")


# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    main()