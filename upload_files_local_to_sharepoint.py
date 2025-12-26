#from scripts_sharepoint import subir_pdfs_en_lote_a_sharepoint, procesar_excel_en_sharepoint_y_limpiar_local,procesar_excel_en_sharepoint_y_limpiar_local_temporal
from scripts_sharepoint import subir_pdfs_en_lote_a_sharepoint, procesar_excel_en_sharepoint_y_limpiar_local,procesar_excel_en_sharepoint_y_limpiar_local_temporal

# Carga de PDFs de APM Local a SharePoint Robótico:
subir_pdfs_en_lote_a_sharepoint(
    local_folder_path=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\APM\Resultado_PDF",
    shared_folder_url="https://unimarcompe.sharepoint.com/:f:/s/UNIMAR-SERVICIOSPORTUARIOS/EhlassOqnRNAng0gYnS-gY8BTU7gbVOlCFow78XHPd-A_g?e=LoKwB7",
    subfolder_path="FACTURAS APMT"
)

# Procesamiento de excel APM Local a SharePoint Robótico:
# Antes: procesar_excel_en_sharepoint_y_limpiar_local
procesar_excel_en_sharepoint_y_limpiar_local_temporal(    
    shared_file_url="https://unimarcompe.sharepoint.com/:x:/s/UNIMAR-SERVICIOSPORTUARIOS/EYawpcM9DyRKt8qiYvA4Rl8BOP4ls5aIKiNpaCm4wzJKyQ?e=T5JafF",
    archivo_local_path=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\APM\Recobro Servicios Portuarios APM 2025.xlsx",
    hoja_objetivo="APMT"
)

# Carga de PDFs de DPW Local a SharePoint Robótico:
subir_pdfs_en_lote_a_sharepoint(
    local_folder_path=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW\Resultado_PDF",
    shared_folder_url="https://unimarcompe.sharepoint.com/:f:/s/UNIMAR-SERVICIOSPORTUARIOS/EryL4-OqqZ5KpnjsS0RAIvcB89XlzlwOQM38EOU2D18eXQ?e=d8w7dc",
    subfolder_path="FACTURAS DPW"
)

# Procesamiento de excel DPW Local a SharePoint Robótico:
procesar_excel_en_sharepoint_y_limpiar_local(
    shared_file_url="https://unimarcompe.sharepoint.com/:x:/s/UNIMAR-SERVICIOSPORTUARIOS/ERXf6KIepFJOlVWRCYOJgkgBKsza4lk0gFmAp2UZXjDrzA?e=cReHcP",
    archivo_local_path=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW\Recobro servicios portuarios DPW 2025.xlsx",
    hoja_objetivo="DPW"
)