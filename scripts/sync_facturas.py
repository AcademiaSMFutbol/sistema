#!/usr/bin/env python3
"""
Sync facturas de proveedores desde correo (IMAP) a Google Drive.
SM Academia — ejecutar localmente o vía tarea programada.

Uso:
  1. Copia config.ejemplo.json → config.json y rellena tus contraseñas
  2. pip install google-api-python-client google-auth
  3. python scripts/sync_facturas.py
"""

import json
import email
import imaplib
import io
import sys
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ── Carpetas Drive ─────────────────────────────────────────────────────────────

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

SUPPLIER_RULES = [
    ('RENTBOX',         ['rentbox', 'rent-box']),
    ('CANVA',           ['canva.com', '@canva']),
    ('ANTHROPIC',       ['anthropic.com']),
    ('MERCADECOR',      ['mercadecor', 'supemegastore']),
    ('BAZAR BUHO',      ['buho', 'búho', 'bazarbuho']),
    ('IMPRENTA ALZOLA', ['alzola']),
    ('DECATHLON',       ['decathlon']),
]

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


def fetch_pdfs(host: str, user: str, password: str, days: int) -> list[dict]:
    results = []
    since = (datetime.now(timezone.utc) - timedelta(days=days)).strftime('%d-%b-%Y')
    try:
        conn = imaplib.IMAP4_SSL(host, 993)
        conn.login(user, password)
        conn.select('INBOX')
        _, msg_ids = conn.search(None, f'SINCE "{since}"')
        ids = msg_ids[0].split() if msg_ids[0] else []
        print(f'  {user}: {len(ids)} mensajes en {days} días')
        for mid in ids:
            _, data = conn.fetch(mid, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])
            sender  = decode_str(msg.get('From', ''))
            subject = decode_str(msg.get('Subject', ''))
            date    = decode_str(msg.get('Date', ''))
            for part in msg.walk():
                ct = part.get_content_type()
                if ct not in ('application/pdf', 'application/octet-stream'):
                    continue
                filename = decode_str(part.get_filename() or '')
                if not filename.lower().endswith('.pdf'):
                    continue
                payload = part.get_payload(decode=True)
                if payload:
                    results.append({
                        'filename': filename,
                        'data':     payload,
                        'sender':   sender,
                        'subject':  subject,
                        'date':     date,
                        'account':  user,
                    })
        conn.logout()
    except Exception as e:
        print(f'  ⚠️  Error en {user}: {e}', file=sys.stderr)
    return results


def build_drive(sa_path: str):
    with open(sa_path) as f:
        info = json.load(f)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def list_files(drive, folder_id: str) -> set[str]:
    names, token = set(), None
    while True:
        r = drive.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields='nextPageToken,files(name)',
            pageToken=token,
        ).execute()
        names.update(f['name'] for f in r.get('files', []))
        token = r.get('nextPageToken')
        if not token:
            break
    return names


def ensure_folder(drive, supplier: str) -> str:
    if supplier in SUPPLIER_FOLDERS:
        return SUPPLIER_FOLDERS[supplier]
    meta = {'name': supplier, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [DRIVE_ROOT]}
    fid = drive.files().create(body=meta, fields='id').execute()['id']
    SUPPLIER_FOLDERS[supplier] = fid
    print(f'  📁 Nueva carpeta: {supplier}')
    return fid


def upload(drive, folder_id: str, filename: str, data: bytes):
    meta  = {'name': filename, 'parents': [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype='application/pdf')
    drive.files().create(body=meta, media_body=media, fields='id').execute()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    config_path = Path(__file__).parent / 'config.json'
    if not config_path.exists():
        print('❌ Falta scripts/config.json — copia config.ejemplo.json y rellena tus datos.')
        sys.exit(1)

    with open(config_path) as f:
        cfg = json.load(f)

    days = cfg.get('dias_atras', 60)

    print(f'\n=== Sync Facturas — {datetime.now().strftime("%Y-%m-%d %H:%M")} ===\n')
    print('Conectando a Drive...')
    drive = build_drive(cfg['google_sa_json'])
    cache = {s: list_files(drive, fid) for s, fid in SUPPLIER_FOLDERS.items()}
    cache['_UNKNOWN'] = list_files(drive, DRIVE_ROOT)
    print(f'Drive listo.\n')

    print('Revisando correos...')
    all_pdfs = []
    for acc in cfg['cuentas']:
        all_pdfs += fetch_pdfs(acc['host'], acc['user'], acc['password'], days)
    print(f'\n{len(all_pdfs)} PDFs encontrados en correo.\n')

    saved, skipped, unknown = [], [], []

    for pdf in all_pdfs:
        supplier = identify_supplier(pdf['sender'], pdf['subject'])
        key      = supplier or '_UNKNOWN'
        folder   = ensure_folder(drive, supplier) if supplier else DRIVE_ROOT

        if pdf['filename'] in cache.get(key, set()):
            skipped.append(pdf['filename'])
            continue

        try:
            upload(drive, folder, pdf['filename'], pdf['data'])
            cache.setdefault(key, set()).add(pdf['filename'])
            saved.append({'supplier': supplier or 'DESCONOCIDO', 'filename': pdf['filename']})
            print(f'  ✅ [{supplier or "DESCONOCIDO"}] {pdf["filename"]}')
        except Exception as e:
            print(f'  ❌ Error subiendo {pdf["filename"]}: {e}', file=sys.stderr)

        if not supplier:
            unknown.append(pdf['filename'])

    print(f'\n{"─"*50}')
    print(f'Guardadas: {len(saved)}  |  Duplicadas omitidas: {len(skipped)}  |  Sin clasificar: {len(unknown)}')
    if unknown:
        print('Sin clasificar (en carpeta raíz — revisa manualmente):')
        for f in unknown:
            print(f'  - {f}')


if __name__ == '__main__':
    main()
