from scripts_sharepoint import descargar_excel_desde_sharepoint

# Descargar Base tipo servicio APM:

descargar_excel_desde_sharepoint(
    shared_file_url="https://unimarcompe.sharepoint.com/:x:/s/UNIMAR-SERVICIOSPORTUARIOS/EX_XTOWeFNNLjYe7LI6abmYB6poMLNTnIkkfdxSG_oZKaA?e=bgNOvk",
    destino_local=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\APM")


# Descargar Base tipo servicio DPW:

descargar_excel_desde_sharepoint(
    shared_file_url="https://unimarcompe.sharepoint.com/:x:/s/UNIMAR-SERVICIOSPORTUARIOS/EWv53zVNQIRKmzBAN5lkNJsBMYAoECEtf_VhQ4aMSvt0OQ?e=Eetwpg",
    destino_local=r"D:\Python\ProcesamientoFaturas\Proyecto - Extraccion y Carga de Facturas APM & DPW\Files\DPW")
