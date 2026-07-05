#!/usr/bin/env python3
"""
Descarga todas las facturas (ingresos + gastos) de Google Drive
y genera un ZIP listo para enviar a la asesoría.

Uso:
  python scripts/descargar_facturas_asesoria.py

Requiere:
  - scripts/config.json con clave "google_sa_json" apuntando al JSON
    de la cuenta de servicio (mismo fichero que usa sync_facturas.py)
  - pip install google-api-python-client google-auth
"""

import io
import json
import zipfile
from pathlib import Path
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ── Carpetas Drive ─────────────────────────────────────────────────────────────

INGRESOS_FOLDER = '1PqKU6DfAElQigDPDkBLnuNEklSZplRBn'   # Facturas emitidas

GASTOS_FOLDERS = {
    'RENTBOX':          '1J81M3_aoXyeQqlFi8gnBRww0ZWYU81j-',
    'CANVA':            '1RI08tmUaXGW5Bm5aOgkhb8Lt3CUw2kVf',
    'ANTHROPIC':        '1NosgtvvIG8-XH2ZT7xRpi7mtfu8fzxwh',
    'MERCADECOR':       '1qEwFZO3aZsD07ueaGoLhxLhc_CN1YU6C',
    'BAZAR BUHO':       '1kB57K7y3Tkzn9Fii7xxOpjsalWQQaHsM',
    'IMPRENTA ALZOLA':  '1bsn7y60-2ntnO0piyLnWhACnUWeywEyh',
    'DECATHLON':        '1I-R31StLHi0EnoKvkYjn0WlDQbxE4cxL',
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def build_drive(sa_path: str):
    with open(sa_path) as f:
        info = json.load(f)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def list_pdfs(drive, folder_id: str) -> list[dict]:
    files, token = [], None
    while True:
        r = drive.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false",
            fields='nextPageToken,files(id,name)',
            pageToken=token,
            pageSize=100,
        ).execute()
        files.extend(r.get('files', []))
        token = r.get('nextPageToken')
        if not token:
            break
    return files


def download_pdf(drive, file_id: str) -> bytes:
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    config_path = Path(__file__).parent / 'config.json'
    if not config_path.exists():
        print('❌ Falta scripts/config.json')
        return

    with open(config_path) as f:
        cfg = json.load(f)

    print('\n=== Descarga Facturas para Asesoría ===\n')
    drive = build_drive(cfg['google_sa_json'])

    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    zip_path = Path(__file__).parent.parent / f'facturas_asesoria_{timestamp}.zip'

    seen_ids: set[str] = set()
    total, errores = 0, 0

    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:

        # ── Facturas emitidas (ingresos) ───────────────────────────────────────
        print('📥 Descargando facturas emitidas (ingresos)...')
        ingresos = list_pdfs(drive, INGRESOS_FOLDER)
        print(f'   {len(ingresos)} archivos encontrados')

        for f in ingresos:
            if f['id'] in seen_ids:
                continue
            seen_ids.add(f['id'])
            try:
                data = download_pdf(drive, f['id'])
                arcname = f'ingresos/{f["name"]}'
                zf.writestr(arcname, data)
                total += 1
                print(f'   ✅ {f["name"]}')
            except Exception as e:
                errores += 1
                print(f'   ❌ {f["name"]}: {e}')

        # ── Facturas de gastos (proveedores) ───────────────────────────────────
        print('\n📥 Descargando facturas de gastos (proveedores)...')
        for proveedor, folder_id in GASTOS_FOLDERS.items():
            archivos = list_pdfs(drive, folder_id)
            if not archivos:
                print(f'   {proveedor}: sin archivos')
                continue
            print(f'   {proveedor}: {len(archivos)} archivos')
            for f in archivos:
                if f['id'] in seen_ids:
                    continue
                seen_ids.add(f['id'])
                try:
                    data = download_pdf(drive, f['id'])
                    arcname = f'gastos/{proveedor}/{f["name"]}'
                    zf.writestr(arcname, data)
                    total += 1
                    print(f'      ✅ {f["name"]}')
                except Exception as e:
                    errores += 1
                    print(f'      ❌ {f["name"]}: {e}')

    print(f'\n{"─"*50}')
    print(f'✅ ZIP generado: {zip_path.name}')
    print(f'   Archivos incluidos: {total}')
    if errores:
        print(f'   Errores: {errores}')
    print(f'   Tamaño: {zip_path.stat().st_size / 1024 / 1024:.1f} MB')


if __name__ == '__main__':
    main()
