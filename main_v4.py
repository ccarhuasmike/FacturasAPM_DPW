from fastapi import FastAPI, Request, UploadFile, File, HTTPException
import shutil
import uvicorn
import os
import logging
from procesamiento_v4 import (
    apm_descomprimir_zip, apm_procesar_pdfs, apm_procesar_excels, apm_limpiar_input,
    dpw_descomprimir_zip, dpw_procesar_pdfs, dpw_procesar_excels, dpw_limpiar_input,
)
from scripts_sharepoint import descargar_y_eliminar_archivos_sharepoint

input_folder_apm = r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\APM\Input"
input_folder_dpw = r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW\Input"

input_folder_apm_update = r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\APM\Input_update"
input_folder_dpw_update = r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW\Input_update"

log_filename = os.path.join(os.path.dirname(__file__), 'procesamiento_log.txt')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
app = FastAPI()

@app.post("/apm_upload")
async def apm_upload_file(request: Request):
    content = await request.body()
    zip_path = os.path.join(input_folder_apm, "archivo_recibido.zip")
    try:
        with open(zip_path, "wb") as f:
            f.write(content)
        logging.info("Archivo ZIP recibido y guardado para APM.")
    except Exception as e:
        logging.error(f"Error guardando archivo ZIP para APM: {e}")
        raise HTTPException(status_code=500, detail="Error guardando archivo ZIP")

    try:
        apm_descomprimir_zip(zip_path)
        logging.info("ZIP descomprimido para APM.")
        resultado_pdf = apm_procesar_pdfs()
        resultado_excel = apm_procesar_excels()
        apm_limpiar_input()
    except Exception as e:
        logging.error(f"Error en el procesamiento de archivos APM: {e}")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento")

    if resultado_pdf and resultado_excel:
        logging.info("Archivos APM procesados exitosamente.")
        return {"message": "Archivos procesados exitosamente"}
    else:
        logging.warning("Hubo errores en el procesamiento de archivos APM.")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento")

@app.post("/dpw_upload")
async def dpw_upload_file(request: Request):
    content = await request.body()
    zip_path = os.path.join(input_folder_dpw, "archivo_recibido.zip")
    try:
        with open(zip_path, "wb") as f:
            f.write(content)
        logging.info("Archivo ZIP recibido y guardado para DPW.")
    except Exception as e:
        logging.error(f"Error guardando archivo ZIP para DPW: {e}")
        raise HTTPException(status_code=500, detail="Error guardando archivo ZIP")

    try:
        dpw_descomprimir_zip(zip_path)
        logging.info("ZIP descomprimido para DPW.")
        resultado_pdf = dpw_procesar_pdfs()
        resultado_excel = dpw_procesar_excels()
        dpw_limpiar_input()
    except Exception as e:
        logging.error(f"Error en el procesamiento de archivos DPW: {e}")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento DPW")

    if resultado_pdf and resultado_excel:
        logging.info("Archivos DPW procesados exitosamente.")
        return {"message": "Archivos DPW procesados exitosamente"}
    else:
        logging.warning("Hubo errores en el procesamiento de archivos DPW.")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento DPW")

@app.post("/dpw_only_files_upload")
async def dpw_only_file_upload():
    try:
        descargar_y_eliminar_archivos_sharepoint(
            shared_folder_url="https://unimarcompe.sharepoint.com/:f:/s/UNIMAR-SERVICIOSPORTUARIOS/EsUC__RmQUNJhVFRy1OM8Q0BJ2ieZMkaYeXRBV_07bRPqg?e=TmQW19",
            carpeta_destino_local=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW\Input"
        )
        logging.info("Archivos descargados desde SharePoint para DPW.")
        resultado_pdf = dpw_procesar_pdfs()
        resultado_excel = dpw_procesar_excels()
        dpw_limpiar_input()
    except Exception as e:
        logging.error(f"Error en el procesamiento DPW sin descomprimir: {e}")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento DPW sin descomprimir")

    if resultado_pdf and resultado_excel:
        logging.info("Archivos DPW procesados exitosamente sin descomprimir ZIP.")
        return {"message": "Archivos DPW procesados exitosamente sin descomprimir ZIP"}
    else:
        logging.warning("Hubo errores en el procesamiento DPW sin descomprimir.")
        raise HTTPException(status_code=500, detail="Hubo errores en el procesamiento DPW sin descomprimir")

#####################################################################################################

if __name__ == "__main__":
    uvicorn.run("main_v4:app", host="0.0.0.0", port=8000, reload=True)