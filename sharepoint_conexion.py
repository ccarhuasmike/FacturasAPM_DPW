import os
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import logging

load_dotenv()

log_filename = os.path.join(os.path.dirname(__file__), 'procesamiento_log.txt')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# ======================= CONFIGURACIÓN ==========================

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ======================= AUTENTICACIÓN ==========================

def get_graph_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Error en autenticación: {result}")


# ======================= FUNCIONES SHAREPOINT ==========================

def get_drive_item_id_from_url(shared_url, token):
    url = f"{GRAPH_BASE}/shares/u!{encode_url_to_share_id(shared_url)}/driveItem"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()


def encode_url_to_share_id(shared_url):
    import base64
    base64url = base64.b64encode(shared_url.encode("utf-8")).decode("utf-8")
    return base64url.replace("/", "_").replace("+", "-").rstrip("=")


# ======================= SUBIDA DE ARCHIVOS ==========================

def upload_file_to_sharepoint(shared_folder_url, local_file_path, remote_filename=None):

    token = get_graph_token()

    # Obtener metadatos del item compartido
    item_data = get_drive_item_id_from_url(shared_folder_url, token)
    drive_id = item_data['parentReference']['driveId']
    raw_path = item_data['parentReference'].get('path', '')

    # Obtener nombre del archivo si no se especificó
    filename = remote_filename or os.path.basename(local_file_path)

    # Extraer solo el path relativo desde '/root:' hacia adelante
    folder_path = ''
    if '/root:' in raw_path:
        folder_path = raw_path.split('/root:')[-1]

    # Construir el path final para la subida
    if folder_path.endswith('/'):
        upload_path = f"{folder_path}{filename}"
    elif folder_path:
        upload_path = f"{folder_path}/{filename}"
    else:
        upload_path = f"/{filename}"

    # Construir la URL de carga en Graph
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{upload_path}:/content"
    logging.info(f"Subiendo archivo a: {url}")

    with open(local_file_path, "rb") as f:
        file_data = f.read()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }

    response = requests.put(url, headers=headers, data=file_data)

    if response.status_code in (200, 201):
        logging.info(f"Archivo subido correctamente: {filename}")
    else:
        logging.error(f"Error al subir archivo: {response.status_code} - {response.text}")
        response.raise_for_status()


def upload_file_to_sharepoint_v2(shared_folder_url, local_file_path, remote_filename=None, subfolder_path=""):

    token = get_graph_token()
    folder_info = get_drive_item_id_from_url(shared_folder_url, token)

    drive_id = folder_info['parentReference']['driveId']
    base_path = folder_info['parentReference'].get('path', '')

    if '/root:' in base_path:
        carpeta_base = base_path.split('/root:')[-1]
    else:
        carpeta_base = ""

    # Combinar con subcarpeta si corresponde
    if subfolder_path:
        if not subfolder_path.startswith("/"):
            subfolder_path = "/" + subfolder_path
        ruta_final = f"{carpeta_base}{subfolder_path}/{os.path.basename(local_file_path)}"
    else:
        ruta_final = f"{carpeta_base}/{os.path.basename(local_file_path)}"

    ruta_final = ruta_final.replace("//", "/")  # evitar rutas dobles

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{ruta_final}:/content"
    headers = {"Authorization": f"Bearer {token}"}

    logging.info(f"Subiendo archivo a: {url}")
    with open(local_file_path, "rb") as f:
        response = requests.put(url, headers=headers, data=f)

    if response.status_code in (200, 201):
        logging.info(f"Subida exitosa: {os.path.basename(local_file_path)}")
    else:
        logging.error(f"Error {response.status_code}: {response.text}")
        response.raise_for_status()

# ======================= DESCARGA DE EXCEL ==========================

def download_file_from_sharepoint(shared_file_url, local_dest_path):
    token = get_graph_token()
    item_data = get_drive_item_id_from_url(shared_file_url, token)

    file_url = f"{GRAPH_BASE}/drives/{item_data['parentReference']['driveId']}/items/{item_data['id']}/content"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(file_url, headers=headers)
    response.raise_for_status()

    with open(local_dest_path, "wb") as f:
        f.write(response.content)

    logging.info(f"Archivo descargado correctamente a: {local_dest_path}")