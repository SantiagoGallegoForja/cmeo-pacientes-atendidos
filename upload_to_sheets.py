"""
Script para subir archivo Excel a Google Sheets
Lee el archivo descargado y lo sube a una hoja de Google Sheets
"""

import os
import json
import logging
import csv
import subprocess
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Configuracion
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def obtener_credenciales():
    """Obtiene las credenciales de Google desde variable de entorno"""
    try:
        creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        if not creds_json:
            raise ValueError("Variable GOOGLE_SHEETS_CREDENTIALS no encontrada")

        creds_dict = json.loads(creds_json)
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        return credentials
    except Exception as e:
        logging.error(f"Error al obtener credenciales: {e}")
        raise

def encontrar_archivos_excel():
    """Encuentra los archivos Excel individuales de reporte (no el combinado)"""
    todos = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'reporte' in f.lower()]
    if not todos:
        raise FileNotFoundError("No se encontro archivo Excel de reporte")

    # Preferir archivos individuales sobre el combinado (que puede estar vacio)
    individuales = [f for f in todos if 'combinado' not in f.lower()]
    archivos = individuales if individuales else todos

    logging.info(f"Archivos encontrados: {archivos}")
    return archivos

def convertir_excel_a_csv(archivo_excel):
    """Convierte Excel a CSV usando LibreOffice/libreoffice (mas robusto)"""
    try:
        csv_file = archivo_excel.replace('.xlsx', '.csv')

        # Intentar con libreoffice (disponible en Ubuntu)
        cmd = [
            'libreoffice', '--headless', '--convert-to', 'csv',
            '--outdir', os.path.dirname(archivo_excel) or '.',
            archivo_excel
        ]

        logging.info(f"Convirtiendo Excel a CSV con LibreOffice...")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

        if result.returncode == 0:
            logging.info(f"Conversion exitosa: {csv_file}")
            return csv_file
        else:
            logging.warning(f"LibreOffice fallo: {result.stderr}")
            raise Exception("LibreOffice conversion failed")

    except Exception as e:
        logging.warning(f"No se pudo convertir con LibreOffice: {e}")
        # Fallback: usar openpyxl con manejo de errores robusto
        return None

def leer_excel_robusto(archivo):
    """Lee Excel de forma robusta, primero intentando CSV"""
    try:
        # Intentar convertir a CSV primero (mas robusto)
        csv_file = convertir_excel_a_csv(archivo)

        if csv_file and os.path.exists(csv_file):
            logging.info(f"Leyendo desde CSV: {csv_file}")
            with open(csv_file, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                data_clean = []
                for row in reader:
                    # Limpiar cada celda
                    clean_row = [str(cell).strip() if cell else "" for cell in row]
                    data_clean.append(clean_row)

            logging.info(f"CSV leido exitosamente: {len(data_clean)} filas")
            return data_clean
        else:
            raise Exception("CSV conversion not available")

    except Exception as e:
        logging.warning(f"No se pudo leer como CSV: {e}")
        logging.info("Intentando lectura directa del Excel...")

        # Fallback: leer Excel byte por byte evitando metadatos
        try:
            from openpyxl import load_workbook
            from openpyxl.utils.exceptions import InvalidFileException

            # Intentar con keep_vba=False y keep_links=False
            wb = load_workbook(
                archivo,
                read_only=True,
                data_only=True,
                keep_vba=False,
                keep_links=False
            )
            ws = wb.active

            logging.info(f"Excel abierto (fallback): {ws.title}")

            data_clean = []
            for row in ws.values:
                clean_row = []
                for cell in row:
                    if cell is None:
                        clean_row.append("")
                    else:
                        clean_row.append(str(cell).strip())
                data_clean.append(clean_row)

            wb.close()
            logging.info(f"Excel leido exitosamente (fallback): {len(data_clean)} filas")
            return data_clean

        except Exception as e2:
            logging.error(f"Error en fallback: {e2}")
            raise

def subir_a_sheets(credentials, sheet_id, data):
    """Sube los datos a Google Sheets"""
    try:
        service = build('sheets', 'v4', credentials=credentials)

        # Nombre de la hoja con fecha de ayer (los datos son de ayer)
        fecha_ayer = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        sheet_name = f"Pacientes_{fecha_ayer}"

        # Crear nueva hoja
        try:
            request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }]
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body=request_body
            ).execute()
            logging.info(f"Nueva hoja creada: {sheet_name}")
        except HttpError as e:
            if "already exists" in str(e):
                logging.info(f"Hoja {sheet_name} ya existe, se actualizara")
            else:
                raise

        # Limpiar hoja existente
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A:ZZ"
        ).execute()

        # Subir datos
        body = {
            'values': data
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()

        logging.info(f"Datos subidos exitosamente: {result.get('updatedCells')} celdas actualizadas")
        return True

    except HttpError as e:
        logging.error(f"Error HTTP al subir a Sheets: {e}")
        raise
    except Exception as e:
        logging.error(f"Error al subir a Sheets: {e}")
        raise

def main():
    """Funcion principal"""
    logging.info("="*60)
    logging.info("SUBIDA A GOOGLE SHEETS")
    logging.info("="*60)

    try:
        # Obtener ID de la hoja
        sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        if not sheet_id:
            raise ValueError("Variable GOOGLE_SHEET_ID no encontrada")

        logging.info(f"Google Sheet ID: {sheet_id}")

        # Obtener credenciales
        credentials = obtener_credenciales()

        # Encontrar y leer archivos Excel (individuales por cuenta)
        archivos = encontrar_archivos_excel()
        data = []
        for archivo in archivos:
            logging.info(f"Procesando: {archivo}")
            # Extraer nombre del doctor del archivo (reporte_pacientes_Daniel.xlsx -> Daniel)
            nombre_base = os.path.splitext(archivo)[0]
            doctor = nombre_base.split('_')[-1] if '_' in nombre_base else "Desconocido"

            archivo_data = leer_excel_robusto(archivo)
            if not data:
                # Primer archivo: agregar columna "Doctor" al encabezado
                if archivo_data:
                    archivo_data[0].append("Doctor")
                    for fila in archivo_data[1:]:
                        fila.append(doctor)
                data = archivo_data
            else:
                # Archivos posteriores: saltar encabezado, agregar columna Doctor
                filas = archivo_data[1:] if len(archivo_data) > 1 else archivo_data
                for fila in filas:
                    fila.append(doctor)
                data.extend(filas)
        logging.info(f"Total filas combinadas: {len(data)}")

        # Subir a Google Sheets
        subir_a_sheets(credentials, sheet_id, data)

        logging.info("="*60)
        logging.info("SUBIDA COMPLETADA EXITOSAMENTE")
        logging.info("="*60)

    except Exception as e:
        logging.error(f"Error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
