#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baixa a planilha do Google Sheets como .xlsx para o build_data.py ler.

- Se houver credenciais de conta de servico no ambiente (GOOGLE_SA_JSON com o
  JSON da chave, ou GOOGLE_APPLICATION_CREDENTIALS com o caminho do arquivo),
  baixa de forma AUTENTICADA via Google Drive API -> a planilha pode ser PRIVADA
  (basta compartilha-la, como Leitor, com o e-mail da conta de servico).
- Sem credenciais, cai no endpoint publico de export (planilha compartilhada
  como "qualquer pessoa com o link").

Requisitos no modo privado:
  * Ativar a **Google Drive API** no projeto do Google Cloud.
  * Bibliotecas: google-api-python-client google-auth
  * (o export de .xlsx pela Drive API tem limite de ~10 MB por arquivo)
"""

import io
import json
import os
import sys
import urllib.request

SHEET_ID = os.environ.get("SHEET_ID", "1U5K5EWqIPMTs6_04NxIPS2DbVXxxkBIIW214aBxBdeI")
OUT = os.environ.get("SRC_XLSX", "BD_-_Conciliacao_Estoque.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def _has_service_account():
    return bool(os.environ.get("GOOGLE_SA_JSON") or
                os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"))


def _validar(data):
    # um .xlsx valido e um zip -> comeca com 'PK'. HTML de login/erro nao.
    if data[:2] != b"PK":
        sys.stderr.write(
            "ERRO: o download nao e um .xlsx valido (veio HTML/pagina de login).\n"
            "  - Modo privado: a planilha foi compartilhada com o e-mail da conta\n"
            "    de servico? A Google Drive API esta ativada no projeto?\n"
            "  - Modo publico: o compartilhamento esta em 'qualquer pessoa com o link'?\n")
        head = data[:300].decode("utf-8", "replace")
        sys.stderr.write("Trecho recebido: " + head + "\n")
        sys.exit(1)


def baixar_autenticado():
    from google.oauth2 import service_account          # lazy import
    from googleapiclient.discovery import build as gbuild
    from googleapiclient.http import MediaIoBaseDownload

    raw = os.environ.get("GOOGLE_SA_JSON")
    if raw:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(raw), scopes=SCOPES)
    else:
        creds = service_account.Credentials.from_service_account_file(
            os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES)

    svc = gbuild("drive", "v3", credentials=creds, cache_discovery=False)
    req = svc.files().export_media(fileId=SHEET_ID, mimeType=XLSX_MIME)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    data = buf.getvalue()
    _validar(data)
    with open(OUT, "wb") as f:
        f.write(data)
    print(f"OK (autenticado) -> {OUT} ({len(data)} bytes)")


def baixar_publico():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        data = r.read()
    _validar(data)
    with open(OUT, "wb") as f:
        f.write(data)
    print(f"OK (publico) -> {OUT} ({len(data)} bytes)")


def main():
    if _has_service_account():
        baixar_autenticado()
    else:
        baixar_publico()


if __name__ == "__main__":
    main()
