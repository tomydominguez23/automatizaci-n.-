"""Sube la planilla a Google Drive y guarda el enlace."""

from __future__ import annotations

import os

from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow

from bot.nube import CLIENTE, SCOPES, TOKEN
from bot.planilla import ROOT

ENV = ROOT / ".env"


def _guardar_id(hoja_id: str) -> None:
    lineas = ENV.read_text(encoding="utf-8").splitlines() if ENV.exists() else []
    nuevas = []
    puesto = False
    for linea in lineas:
        if linea.startswith("GOOGLE_SHEETS_ID="):
            nuevas.append(f"GOOGLE_SHEETS_ID={hoja_id}")
            puesto = True
        else:
            nuevas.append(linea)
    if not puesto:
        nuevas.append(f"GOOGLE_SHEETS_ID={hoja_id}")
    ENV.write_text("\n".join(nuevas) + "\n", encoding="utf-8")


def main() -> None:
    load_dotenv(ENV)
    if not CLIENTE.exists():
        raise SystemExit(
            "Falta credenciales/cliente.json.\n"
            "1. Entra a https://console.cloud.google.com/apis/credentials\n"
            "2. Crea un proyecto y activa Google Sheets API y Google Drive API.\n"
            "3. Configura la pantalla de consentimiento (usuario de prueba: tu Gmail).\n"
            "4. Crea un ID de cliente de tipo Aplicación de escritorio.\n"
            "5. Descarga el JSON y guárdalo como credenciales/cliente.json\n"
            "6. Vuelve a correr: python3 -m bot.vincular"
        )
    TOKEN.parent.mkdir(parents=True, exist_ok=True)
    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENTE), list(SCOPES))
    cred = flow.run_local_server(port=0, open_browser=True)
    TOKEN.write_text(cred.to_json(), encoding="utf-8")

    hoja_id = os.environ.get("GOOGLE_SHEETS_ID", "").strip()
    if not hoja_id:
        raise SystemExit("Falta GOOGLE_SHEETS_ID en .env")
    import gspread

    libro = gspread.authorize(cred).open_by_key(hoja_id)
    print("Conectado a:", libro.title)
    print("https://docs.google.com/spreadsheets/d/" + hoja_id)
    _guardar_id(hoja_id)


if __name__ == "__main__":
    main()
