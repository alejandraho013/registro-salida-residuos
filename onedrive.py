"""
onedrive.py – Integración Microsoft Graph API para TINTATEX
Versión: Conexión a Excel en la nube (Sin Fotos)
"""

import io
import requests
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────
# 1. AUTENTICACIÓN (Token de Microsoft)
# ─────────────────────────────────────────────────────────────

@st.cache_resource(ttl=3300)
def _get_access_token() -> str:
    """Obtiene el token de Azure AD usando las credenciales de los Secrets."""
    import msal
    tenant_id     = st.secrets["AZURE_TENANT_ID"]
    client_id     = st.secrets["AZURE_CLIENT_ID"]
    client_secret = st.secrets["AZURE_CLIENT_SECRET"]

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    # Buscamos el permiso para Microsoft Graph
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" not in result:
        raise RuntimeError(f"Error de conexión con Microsoft: {result.get('error_description')}")
    return result["access_token"]

def _headers() -> dict:
    """Cabeceras estándar para las peticiones a la API."""
    return {
        "Authorization": f"Bearer {_get_access_token()}",
        "Content-Type": "application/json"
    }

GRAPH = "https://graph.microsoft.com/v1.0"

def _file_url() -> str:
    """Construye la URL del archivo usando el ID guardado en Secrets."""
    file_id = st.secrets["ONEDRIVE_FILE_ID"]
    return f"{GRAPH}/me/drive/items/{file_id}"

# ─────────────────────────────────────────────────────────────
# 2. FUNCIONES DE LECTURA Y BÚSQUEDA
# ─────────────────────────────────────────────────────────────

def get_file_id(filename: str = "database.xlsx") -> str:
    """Busca el archivo en OneDrive y devuelve su ID único."""
    url = f"{GRAPH}/me/drive/root/search(q='{filename}')"
    resp = requests.get(url, headers=_headers())
    resp.raise_for_status()
    items = resp.json().get("value", [])
    if not items:
        raise FileNotFoundError(f"No se encontró el archivo '{filename}' en la raíz de OneDrive.")
    return items[0]["id"]

@st.cache_data(ttl=120)
def cargar_datos_onedrive() -> pd.DataFrame:
    """Descarga el Excel de OneDrive y lee la hoja MASTER."""
    try:
        url = _file_url() + "/content"
        resp = requests.get(url, headers=_headers())
        resp.raise_for_status()

        df = pd.read_excel(io.BytesIO(resp.content), sheet_name="MASTER", engine="openpyxl")
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        return df
    except Exception as e:
        st.error(f"❌ Error al cargar desde OneDrive: {e}")
        return pd.DataFrame()

# ─────────────────────────────────────────────────────────────
# 3. ESCRITURA (Append de filas al final)
# ─────────────────────────────────────────────────────────────

def append_filas_onedrive(nuevas_filas: list[dict]) -> None:
    """Agrega registros al final del Excel sin borrar lo anterior."""
    base_url = _file_url()
    hdrs = _headers()

    # Abrir una sesión de trabajo en el libro
    sess_resp = requests.post(base_url + "/workbook/createSession", headers=hdrs, json={"persistChanges": True})
    sess_resp.raise_for_status()
    session_id = sess_resp.json()["id"]
    hdrs["workbook-session-id"] = session_id

    try:
        # 1. Identificar el rango usado actualmente para saber dónde empezar a escribir
        used_resp = requests.get(base_url + "/workbook/worksheets/MASTER/usedRange", headers=hdrs)
        used_resp.raise_for_status()
        # El número de filas actual nos dice la posición de inicio
        n_filas_actuales = len(used_resp.json()["values"])
        
        # 2. Definir las columnas exactas (deben coincidir con el orden del Excel)
        COLUMNAS = ["fecha", "mes", "empresa", "conductor", "placa", "tipo_residuo", "peso_kg", "novedades"]
        
        # 3. Preparar los datos
        valores = [[str(row.get(c, "")) for c in COLUMNAS] for row in nuevas_filas]

        # 4. Calcular el rango (Desde A hasta H)
        # Ejemplo: Si hay 10 filas, escribimos desde la A11 hasta la H(11 + registros nuevos)
        rango = f"A{n_filas_actuales + 1}:H{n_filas_actuales + len(valores)}"

        # 5. Insertar los datos
        patch_resp = requests.patch(
            base_url + f"/workbook/worksheets/MASTER/range(address='{rango}')",
            headers=hdrs,
            json={"values": valores}
        )
        patch_resp.raise_for_status()

    finally:
        # Siempre cerrar la sesión para liberar el archivo
        requests.post(base_url + "/workbook/closeSession", headers=hdrs)

def descargar_excel_onedrive() -> bytes:
    """Descarga el archivo completo para el botón de exportación."""
    url = _file_url() + "/content"
    resp = requests.get(url, headers=_headers())
    resp.raise_for_status()
    return resp.content
