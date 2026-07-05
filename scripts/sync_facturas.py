#!/usr/bin/env python3
"""
Sync facturas de proveedores desde correo (IMAP) a Google Drive.
SM Academia — ejecutado por GitHub Actions cada día.

Requiere los siguientes GitHub Secrets:
  GMAIL_USER, GMAIL_APP_PASSWORD
  OUTLOOK_ACCOUNTS  → JSON: [{"user":"...","password":"..."}]
  CUSTOM_ACCOUNTS   → JSON: [{"host":"...","user":"...","password":"..."}]
  GOOGLE_SA_JSON    → JSON de la cuenta de servicio de Google Drive
"""

import os
import sys
import json
import email
import imaplib
import io
from datetime import datetime, timedelta, timezone
from email.header import decode_header

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ── Configuración Drive ────────────────────────────────────────────────────────

DRIVE_ROOT = '1i09bvwehgbDlaywAu1ED_M0J-V_j3Oqv'

SUPPLIER_FOLDERS = {
    'RENTBOX':          '1J81M3_aoXyeQqlFi8gnBRww0ZWYU81j-',
    'CANVA':            '1RI08tmUaXGW5Bm5aOgkhb8Lt3CUw2kVf',
    'ANTHROPIC':        '1NosgtvvIG8-XH2ZT7xRpi7mtfu8fzxwh',
    'MERCADECOR':       '1qEwFZO3aZsD07ueaGoLhxLhc_CN1YU6C',
    'BAZAR BUHO':       '1kB57K7y3Tkzn9Fii7xxOpjsalWQQaHsM',
    'IMPRENTA ALZOLA':  '1bsn7y60-2ntnO0piyLnWhACnUWeywEyh',
    'DECATHLON':        '1I-R31StLHi0EnoKvkYjn0WlDQbxE4cxL',
}

# (proveedor, [palabras clave en remitente o asunto])
SUPPLIER_RULES = [
    ('RENTBOX',         ['rentbox', 'rent-box']),
    ('CANVA',           ['canva.com', 'canva']),
    ('ANTHROPIC',       ['anthropic.com']),
    ('MERCADECOR',      ['mercadecor', 'supemegastore']),
    ('BAZAR BUHO',      ['buho', 'búho', 'bazarbuho', 'bazar buho']),
    ('IMPRENTA ALZOLA', ['alzola']),
    ('DECATHLON',       ['decathlon']),
]

IMAP_SEARCH_DAYS = 60

# ── Helpers ────────────────────────────────────────────────────────────────────

def decode_str(value):
    if not value:
        return ''
    parts = decode_header(value)
    out = []
    for raw, enc in parts:
        if isinstance(raw, bytes):
            out.append(raw.decode(enc or 'utf-8', errors='replace'))
        else:
            out.append(raw)
    return ' '.join(out)


def identify_supplier(sender: str, subject: str) -> str | None:
    text = (sender + ' ' + subject).lower()
    for supplier, keywords in SUPPLIER_RULES:
        if any(kw in text for kw in keywords):
            return supplier
    return None


def imap_since_date() -> str:
    d = datetime.now(timezone.utc) - timedelta(days=IMAP_SEARCH_DAYS)
    return d.strftime('%d-%b-%Y')


def fetch_pdfs_from_imap(host: str, user: str, password: str) -> list[dict]:
    """Conecta por IMAP y devuelve lista de {filename, data, sender, subject, date}."""
    results = []
    try:
        conn = imaplib.IMAP4_SSL(host, 993)
        conn.login(user, password)
        conn.select('INBOX')

        since = imap_since_date()
        _, msg_ids = conn.search(None, f'(SINCE "{since}" HAS X-GM-RAW "has:attachment")')
        # Fallback para servidores que no soporten X-GM-RAW
        if not msg_ids or not msg_ids[0]:
            _, msg_ids = conn.search(None, f'SINCE "{since}"')

        ids = msg_ids[0].split() if msg_ids[0] else []
        print(f'  [{user}] {len(ids)} mensajes desde hace {IMAP_SEARCH_DAYS} días')

        for mid in ids:
            _, data = conn.fetch(mid, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])
            sender  = decode_str(msg.get('From', ''))
            subject = decode_str(msg.get('Subject', ''))
            date    = decode_str(msg.get('Date', ''))

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                if part.get_content_type() not in ('application/pdf', 'application/octet-stream'):
                    continue
                filename = decode_str(part.get_filename() or '')
                if not filename.lower().endswith('.pdf'):
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                results.append({
                    'filename': filename,
                    'data':     payload,
                    'sender':   sender,
                    'subject':  subject,
                    'date':     date,
                    'account':  user,
                })

        conn.logout()
    except imaplib.IMAP4.error as e:
        print(f'  ⚠️  Error IMAP [{user}@{host}]: {e}', file=sys.stderr)
    except Exception as e:
        print(f'  ⚠️  Error inesperado [{user}@{host}]: {e}', file=sys.stderr)
    return results


# ── Drive ──────────────────────────────────────────────────────────────────────

def build_drive_service():
    sa_json = os.environ.get('GOOGLE_SA_JSON', '')
    if not sa_json:
        raise RuntimeError('GOOGLE_SA_JSON no configurado')
    info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def list_existing_files(drive, folder_id: str) -> set[str]:
    existing = set()
    page_token = None
    while True:
        resp = drive.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields='nextPageToken, files(name)',
            pageToken=page_token,
        ).execute()
        for f in resp.get('files', []):
            existing.add(f['name'])
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return existing


def ensure_supplier_folder(drive, supplier: str) -> str:
    if supplier in SUPPLIER_FOLDERS:
        return SUPPLIER_FOLDERS[supplier]
    # Crear subcarpeta nueva
    meta = {
        'name': supplier,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [DRIVE_ROOT],
    }
    folder = drive.files().create(body=meta, fields='id').execute()
    fid = folder['id']
    SUPPLIER_FOLDERS[supplier] = fid
    print(f'  📁 Nueva carpeta creada para {supplier}: {fid}')
    return fid


def upload_to_drive(drive, folder_id: str, filename: str, data: bytes):
    meta = {'name': filename, 'parents': [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype='application/pdf', resumable=False)
    drive.files().create(body=meta, media_body=media, fields='id').execute()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f'\n=== Sync Facturas Proveedores — {datetime.now().strftime("%Y-%m-%d %H:%M")} ===\n')

    # 1. Construir servicio Drive y cachear archivos existentes
    print('Conectando a Google Drive...')
    drive = build_drive_service()
    existing_cache: dict[str, set[str]] = {}
    for supplier, fid in SUPPLIER_FOLDERS.items():
        existing_cache[supplier] = list_existing_files(drive, fid)
    existing_cache['_UNKNOWN'] = list_existing_files(drive, DRIVE_ROOT)
    print(f'  Drive listo. {sum(len(v) for v in existing_cache.values())} archivos indexados.\n')

    # 2. Recopilar todos los PDFs de correo
    all_pdfs: list[dict] = []

    # Gmail
    gmail_user = os.environ.get('GMAIL_USER', '')
    gmail_pass = os.environ.get('GMAIL_APP_PASSWORD', '')
    if gmail_user and gmail_pass:
        print(f'📧 Gmail: {gmail_user}')
        all_pdfs += fetch_pdfs_from_imap('imap.gmail.com', gmail_user, gmail_pass)
    else:
        print('⚠️  GMAIL_USER / GMAIL_APP_PASSWORD no configurados')

    # Outlook (outlook.es, hotmail.es, etc.)
    outlook_raw = os.environ.get('OUTLOOK_ACCOUNTS', '[]')
    for acc in json.loads(outlook_raw):
        print(f'📧 Outlook: {acc["user"]}')
        all_pdfs += fetch_pdfs_from_imap('imap-mail.outlook.com', acc['user'], acc['password'])

    # Cuentas dominio propio (direccion@sm-academia.com, samy.martin@academiasmfutbol.com)
    custom_raw = os.environ.get('CUSTOM_ACCOUNTS', '[]')
    for acc in json.loads(custom_raw):
        print(f'📧 Custom [{acc["host"]}]: {acc["user"]}')
        all_pdfs += fetch_pdfs_from_imap(acc['host'], acc['user'], acc['password'])

    print(f'\nTotal PDFs encontrados en correo: {len(all_pdfs)}\n')

    # 3. Clasificar y subir
    saved, skipped, unknown = [], [], []

    for pdf in all_pdfs:
        supplier = identify_supplier(pdf['sender'], pdf['subject'])
        filename = pdf['filename']

        if supplier:
            folder_id = ensure_supplier_folder(drive, supplier)
            cache_key = supplier
        else:
            folder_id = DRIVE_ROOT
            cache_key = '_UNKNOWN'

        if filename in existing_cache.get(cache_key, set()):
            skipped.append({'supplier': supplier or 'DESCONOCIDO', 'filename': filename})
            continue

        try:
            upload_to_drive(drive, folder_id, filename, pdf['data'])
            existing_cache.setdefault(cache_key, set()).add(filename)
            entry = {
                'supplier': supplier or 'DESCONOCIDO',
                'filename': filename,
                'date':     pdf['date'],
                'account':  pdf['account'],
            }
            saved.append(entry)
            print(f'  ✅ {supplier or "DESCONOCIDO"} → {filename}')
        except Exception as e:
            print(f'  ❌ Error subiendo {filename}: {e}', file=sys.stderr)

        if not supplier:
            unknown.append(filename)

    # 4. Resumen
    print(f'\n{"="*60}')
    print(f'RESUMEN')
    print(f'{"="*60}')
    print(f'✅ Facturas guardadas:  {len(saved)}')
    print(f'⏭️  Duplicadas omitidas: {len(skipped)}')
    print(f'❓ Proveedor desconocido: {len(unknown)}')

    if saved:
        print('\nGuardadas:')
        for s in saved:
            print(f'  [{s["supplier"]}] {s["filename"]} ({s["account"]})')

    if unknown:
        print('\nProveedor desconocido (en carpeta raíz, clasificar manualmente):')
        for f in unknown:
            print(f'  - {f}')

    print()


if __name__ == '__main__':
    main()
