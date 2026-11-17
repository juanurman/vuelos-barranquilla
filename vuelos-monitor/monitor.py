import asyncio
import os
import re
import sys
import smtplib
import subprocess
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import zoneinfo
import openpyxl
from playwright.async_api import async_playwright

EXCEL_FILE = os.environ.get("EXCEL_FILE", "vuelos_bue_barranquilla_nov2026.xlsx")
GDRIVE_FILE_ID = os.environ.get("GDRIVE_FILE_ID", "1XxKOfrundPS-6D5d73OSK90gXIruaYrp")
ALERT_EMAIL = os.environ.get("ALERT_EMAIL", "urmandiego@gmail.com")
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
ALERT_THRESHOLD = float(os.environ.get("ALERT_THRESHOLD", "1500"))

ROUTES = [
    {
        "key": "BUE-BAQ_20_22",
        "name": "BUE -> BAQ (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BAQ%20from%20BUE%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "MVD-BAQ_20_22",
        "name": "MVD -> BAQ (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BAQ%20from%20MVD%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "BUE-CTG_20_22",
        "name": "BUE -> CTG (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20CTG%20from%20BUE%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "BUE-BOG_20_22",
        "name": "BUE -> BOG (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BOG%20from%20BUE%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "BOG-BAQ_20_22",
        "name": "BOG -> BAQ (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BAQ%20from%20BOG%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "BUE-LIM_20_22",
        "name": "BUE -> LIM (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20LIM%20from%20BUE%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "LIM-BAQ_20_22",
        "name": "LIM -> BAQ (20-22 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BAQ%20from%20LIM%20on%202026-11-20%20through%202026-11-22&curr=USD",
    },
    {
        "key": "BUE-BAQ_19_23",
        "name": "BUE -> BAQ (19-23 nov)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BAQ%20from%20BUE%20on%202026-11-19%20through%202026-11-23&curr=USD",
    },
]

async def scrape_single_route(context, route_info):
    url = route_info["url"]
    name = route_info["name"]
    page = await context.new_page()
    try:
        await page.goto(url, wait_until="networkidle", timeout=35000)
        await page.wait_for_timeout(1500)
        
        btn = page.get_by_role("button", name=re.compile(r"Proceed anyway|Continuar|Aceptar|Accept all", re.I))
        if await btn.count() > 0:
            try:
                await btn.first.click()
                await page.wait_for_timeout(2500)
            except Exception:
                pass
        
        text = await page.inner_text("body")
        
        m = re.search(r"(?:from|desde)\s*(?:USD|US\$|\$)\s*([\d,]+)", text, re.I)
        if m:
            val = int(m.group(1).replace(",", ""))
            await page.close()
            return val
        
        matches = re.findall(r"(?:USD|US\$|\$)\s*([\d,]+)", text)
        valid_prices = []
        for p in matches:
            v = int(p.replace(",", ""))
            if 50 <= v <= 15000:
                valid_prices.append(v)
        
        await page.close()
        if valid_prices:
            return min(valid_prices)
        return "Sin dato"
    except Exception as e:
        print(f"[{name}] Advertencia de extracción: {e}", file=sys.stderr)
        try:
            await page.close()
        except Exception:
            pass
        return "Sin dato"

async def fetch_all_prices():
    results = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-setuid-sandbox",
            ]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
            locale="es-419",
            ignore_https_errors=True
        )
        for r in ROUTES:
            price = await scrape_single_route(context, r)
            results[r["key"]] = price
            print(f"  {r['name']}: {price}")
            await asyncio.sleep(1.0)
        await browser.close()
    return results

def get_arg_time():
    try:
        tz = zoneinfo.ZoneInfo("America/Argentina/Buenos_Aires")
        return datetime.now(tz)
    except Exception:
        return datetime.utcnow()

def update_excel(prices, now_dt):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet_hist = wb["Historial Monitoreo"]
    new_row = sheet_hist.max_row + 1
    
    ts_str = now_dt.strftime("%Y-%m-%d %H:%M")
    
    col_vals = [
        ts_str,
        prices.get("BUE-BAQ_20_22", "Sin dato"),
        prices.get("MVD-BAQ_20_22", "Sin dato"),
        prices.get("BUE-CTG_20_22", "Sin dato"),
        prices.get("BUE-BOG_20_22", "Sin dato"),
        prices.get("BOG-BAQ_20_22", "Sin dato"),
        prices.get("BUE-LIM_20_22", "Sin dato"),
        prices.get("LIM-BAQ_20_22", "Sin dato"),
        prices.get("BUE-BAQ_19_23", "Sin dato"),
        f'=IF(AND(ISNUMBER(G{new_row}),ISNUMBER(H{new_row})),G{new_row}+H{new_row},"-")',
        f'=IF(AND(ISNUMBER(E{new_row}),ISNUMBER(F{new_row})),E{new_row}+F{new_row},"-")',
        f'=MIN(B{new_row},C{new_row},D{new_row},I{new_row})'
    ]
    
    for c_idx, val in enumerate(col_vals, start=1):
        sheet_hist.cell(row=new_row, column=c_idx, value=val)
    
    if "Resumen Comparativo" in wb.sheetnames:
        sheet_res = wb["Resumen Comparativo"]
        sheet_res.cell(row=4, column=1, value=f"Último chequeo automático: {ts_str}")
        bue_p = prices.get("BUE-BAQ_20_22")
        mvd_p = prices.get("MVD-BAQ_20_22")
        bue_str = f"US$ {bue_p:,.0f}".replace(",", ".") if isinstance(bue_p, (int, float)) else str(bue_p)
        mvd_str = f"US$ {mvd_p:,.0f}".replace(",", ".") if isinstance(mvd_p, (int, float)) else str(mvd_p)
        sheet_res.cell(row=4, column=3, value=f"BUE-BAQ 20-22 mín actual: {bue_str}")
        sheet_res.cell(row=4, column=5, value=f"MVD-BAQ mín actual: {mvd_str}")

    wb.save(EXCEL_FILE)
    print(f"Excel guardado exitosamente en fila {new_row}.")
    return new_row

def update_google_drive():
    gdrive_bin = "/opt/antigravity/bin/gdrive"
    if os.path.exists(gdrive_bin):
        try:
            cmd = [gdrive_bin, "mutate", "update", GDRIVE_FILE_ID, EXCEL_FILE]
            res = subprocess.run(cmd, capture_output=True, text=True, check=True)
            print(f"Google Drive actualizado vía gdrive CLI: {res.stdout.strip()}")
            return True
        except Exception as e:
            print(f"Nota CLI gdrive: {e}", file=sys.stderr)
    
    sa_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaFileUpload
            import json
            
            creds_info = json.loads(sa_json)
            creds = service_account.Credentials.from_service_account_info(
                creds_info, scopes=["https://www.googleapis.com/auth/drive"]
            )
            service = build("drive", "v3", credentials=creds)
            media = MediaFileUpload(EXCEL_FILE, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            service.files().update(fileId=GDRIVE_FILE_ID, media_body=media).execute()
            print("Google Drive actualizado vía API oficial.")
            return True
        except Exception as e:
            print(f"Error actualizando Google Drive vía API: {e}", file=sys.stderr)
    
    return False

def send_email(subject, html_body, attach_file=False):
    if not SMTP_USER or not SMTP_PASS:
        print(f"[NOTIFICACIÓN LOCAL] Asunto: {subject}")
        print(f"Destinatario: {ALERT_EMAIL}")
        return False
    
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = ALERT_EMAIL
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html"))
    
    if attach_file and os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(EXCEL_FILE)}")
        msg.attach(part)
    
    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
        server.quit()
        print(f"Correo enviado exitosamente a {ALERT_EMAIL}: {subject}")
        return True
    except Exception as e:
        print(f"Error enviando correo a {ALERT_EMAIL}: {e}", file=sys.stderr)
        return False

def check_immediate_alerts(prices, now_dt):
    alerts = []
    
    bog_combo = None
    if isinstance(prices.get("BUE-BOG_20_22"), (int, float)) and isinstance(prices.get("BOG-BAQ_20_22"), (int, float)):
        bog_combo = prices["BUE-BOG_20_22"] + prices["BOG-BAQ_20_22"]
        
    lim_combo = None
    if isinstance(prices.get("BUE-LIM_20_22"), (int, float)) and isinstance(prices.get("LIM-BAQ_20_22"), (int, float)):
        lim_combo = prices["BUE-LIM_20_22"] + prices["LIM-BAQ_20_22"]

    for r in ROUTES:
        val = prices.get(r["key"])
        if isinstance(val, (int, float)) and val < ALERT_THRESHOLD:
            alerts.append((r["name"], val))
            
    if bog_combo and bog_combo < ALERT_THRESHOLD:
        alerts.append(("Combo Bogotá (BUE->BOG + BOG->BAQ)", bog_combo))
    if lim_combo and lim_combo < ALERT_THRESHOLD:
        alerts.append(("Combo Lima (BUE->LIM + LIM->BAQ)", lim_combo))

    if alerts:
        print(f"🚨 ALERTA: {len(alerts)} opciones por debajo de US$ {ALERT_THRESHOLD:,.0f}!")
        subject = f"🚨 ALERTA: Tarifa encontrada por debajo de US$ {ALERT_THRESHOLD:,.0f} para Barranquilla!"
        rows_html = "".join([f"<tr><td style='padding:8px;border:1px solid #ddd;'><b>{name}</b></td><td style='padding:8px;border:1px solid #ddd;color:#2e7d32;font-size:16px;'><b>US$ {price:,.0f}</b></td></tr>" for name, price in alerts])
        body = f"""
        <h2>¡Alerta de Vuelo Barato a Barranquilla!</h2>
        <p>Se detectaron tarifas por debajo del umbral objetivo de <b>US$ {ALERT_THRESHOLD:,.0f}</b> en el chequeo de las {now_dt.strftime('%H:%M ARG')}:</p>
        <table style='border-collapse:collapse;width:100%;max-width:600px;'>
            <thead>
                <tr style='background:#f2f2f2;'><th style='padding:8px;border:1px solid #ddd;'>Ruta / Estrategia</th><th style='padding:8px;border:1px solid #ddd;'>Precio (USD)</th></tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        <p style='margin-top:20px;'><a href='https://drive.google.com/file/d/{GDRIVE_FILE_ID}/view' style='background:#1a73e8;color:#fff;padding:10px 18px;text-decoration:none;border-radius:4px;'>Abrir Excel en Google Drive</a></p>
        """
        send_email(subject, body, attach_file=True)

def check_daily_report(prices, now_dt, force=False):
    if now_dt.hour == 20 or force:
        print("Enviando Resumen Diario de las 20:00...")
        subject = f"📊 Resumen Diario Monitoreo Vuelos BUE -> BAQ ({now_dt.strftime('%d/%m/%Y')} 20:00 ARG)"
        
        table_rows = []
        for r in ROUTES:
            p = prices.get(r["key"], "Sin dato")
            p_str = f"US$ {p:,.0f}" if isinstance(p, (int, float)) else str(p)
            table_rows.append(f"<tr><td style='padding:6px 10px;border:1px solid #ddd;'>{r['name']}</td><td style='padding:6px 10px;border:1px solid #ddd;font-weight:bold;'>{p_str}</td></tr>")
        
        bog_combo = "-"
        if isinstance(prices.get("BUE-BOG_20_22"), (int, float)) and isinstance(prices.get("BOG-BAQ_20_22"), (int, float)):
            bog_combo = f"US$ {prices['BUE-BOG_20_22'] + prices['BOG-BAQ_20_22']:,.0f}"
        table_rows.append(f"<tr style='background:#f9f9f9;'><td style='padding:6px 10px;border:1px solid #ddd;'>Combo Bogotá (BUE->BOG + BOG->BAQ)</td><td style='padding:6px 10px;border:1px solid #ddd;font-weight:bold;'>{bog_combo}</td></tr>")

        lim_combo = "-"
        if isinstance(prices.get("BUE-LIM_20_22"), (int, float)) and isinstance(prices.get("LIM-BAQ_20_22"), (int, float)):
            lim_combo = f"US$ {prices['BUE-LIM_20_22'] + prices['LIM-BAQ_20_22']:,.0f}"
        table_rows.append(f"<tr style='background:#f9f9f9;'><td style='padding:6px 10px;border:1px solid #ddd;'>Combo Lima (BUE->LIM + LIM->BAQ)</td><td style='padding:6px 10px;border:1px solid #ddd;font-weight:bold;'>{lim_combo}</td></tr>")

        body = f"""
        <h2>Resumen Diario de Monitoreo - Vuelos a Barranquilla</h2>
        <p>Fecha y Hora: <b>{now_dt.strftime('%d/%m/%Y %H:%M')} ARG</b></p>
        <p>Estado actual de las tarifas registradas:</p>
        <table style='border-collapse:collapse;width:100%;max-width:650px;'>
            <thead>
                <tr style='background:#1a73e8;color:#fff;'><th style='padding:8px 10px;border:1px solid #ddd;text-align:left;'>Ruta</th><th style='padding:8px 10px;border:1px solid #ddd;text-align:left;'>Mínimo Observado</th></tr>
            </thead>
            <tbody>
                {''.join(table_rows)}
            </tbody>
        </table>
        <p style='margin-top:20px;'>Adjunto encontrarás la versión completa y actualizada del archivo Excel.</p>
        <p><a href='https://drive.google.com/file/d/{GDRIVE_FILE_ID}/view'>Ver archivo en Google Drive</a></p>
        """
        send_email(subject, body, attach_file=True)

async def run_once(force_daily=False):
    now_dt = get_arg_time()
    print(f"\n[{now_dt.strftime('%Y-%m-%d %H:%M:%S ARG')}] Chequeo horario iniciado...")
    prices = await fetch_all_prices()
    update_excel(prices, now_dt)
    update_google_drive()
    check_immediate_alerts(prices, now_dt)
    check_daily_report(prices, now_dt, force=force_daily)
    print(f"[{now_dt.strftime('%H:%M:%S ARG')}] Chequeo completado.")

async def main():
    force_daily = "--daily-report" in sys.argv
    loop_mode = "--loop" in sys.argv
    
    if loop_mode:
        print("Iniciando modo bucle en segundo plano (cada 1 hora)...")
        while True:
            await run_once(force_daily=force_daily)
            print("Esperando 3600 segundos (1 hora) para el próximo chequeo...")
            await asyncio.sleep(3600)
    else:
        await run_once(force_daily=force_daily)

if __name__ == "__main__":
    asyncio.run(main())
