import pandas as pd
import numpy as np
import win32com.client as win32
import datetime
from tzlocal import get_localzone





# Importar la librería numpy para manejar arrays y operaciones numéricas
# Importar la librería pandas para manejar DataFrames
# Importar la librería openpyxl para trabajar con archivos Excel



# Definimos nombre del archivo de origen y destino
informe_mtto = r'\\S31889004.services.dekra.com\DKI\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\LearningReportAlerts\informe_mtto.xlsx'
informe_destino = r'\\S31889004.services.dekra.com\DKI\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\LearningReportAlerts\ExcelPadre.xlsx'
# Definimos nombre de las hojas de origen y destino
hoja_mtto = 'Hoja1' #Nombre de la hoja del archivo de origen
hoja_destino = 'Lidar Windcube' #Nombre de la hoja del archivo destino
hoja_destino_historico = 'Historico' #Nombre de la hoja del archivo destino para el historico



# Importar las funciones de mantenimiento desde el módulo mantenimiento.py
# Asegúrate de que el archivo mantenimiento.py esté en el mismo directorio o en el PYTHONPATH
from mantenimiento_LRA import (leer_datos_origen, actualizar_fecha_destino, actualizar_comentario_destino, actualizar_metanol_destino, actualizar_liquido_destino, \
                                actualizar_filtros_destino, actualizar_escobilla_destino, actualizar_incidencias_destino, actualizar_baterias_destino, actualizar_sensores_destino,\
                                encontrar_fila_destino, encontrar_fila_historico, leer_correos_origen
                            )


import os
import pandas as pd
import win32com.client as win32
from tzlocal import get_localzone  # zoneinfo de tu Windows

def enviar_correo(destinatario, asunto, mensaje, hora_envio=None, cc=None, adjuntos=None):
    """
    Programar el envío de un correo usando Outlook instalado en Windows.
    
    Parámetros:
    - destinatario: str o lista de correos
    - asunto: str
    - mensaje: str (texto plano)
    - hora_envio: datetime.datetime o pandas.Timestamp (opcional, hora programada)
    - cc: str o lista de correos (opcional)
    - adjuntos: lista de rutas de archivos (opcional)
    """
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0 = MailItem

    # Destinatarios
    mail.To = ";".join(destinatario) if isinstance(destinatario, list) else destinatario

    # CC
    if cc:
        mail.CC = ";".join(cc) if isinstance(cc, list) else cc

    # Asunto y cuerpo
    mail.Subject = asunto
    mail.Body = mensaje

    # Adjuntos
    if adjuntos:
        for archivo in adjuntos:
            if os.path.isfile(archivo):
                mail.Attachments.Add(archivo)

    # Guardar los destinatarios antes de enviar
    destinatarios_str = mail.To

    # Programar envío
    if hora_envio:
        if isinstance(hora_envio, pd.Timestamp):
            hora_envio = hora_envio.to_pydatetime()

        # Si es naive, añadirle la zona local de Windows
        if hora_envio.tzinfo is None:
            local_tz = get_localzone()  # devuelve un ZoneInfo
            hora_envio = hora_envio.replace(tzinfo=local_tz)

        mail.DeferredDeliveryTime = hora_envio

    mail.Send()
    print(f"✅ Correo se enviará a {destinatarios_str}" + (f" a las {hora_envio}" if hora_envio else ""))








def main():
    # Leer datos del archivo de origen:
    n_equipo, fecha_ultimomantenimiento, comentario_ultimomantenimiento, metanol_ultimomantenimiento, liquido_ultimomantenimiento, codigo_incidencias , sensores_cambiados_ultimomantenimiento, estado_baterias = leer_datos_origen(informe_mtto, hoja_mtto)
    
    print(n_equipo, fecha_ultimomantenimiento, comentario_ultimomantenimiento, metanol_ultimomantenimiento, liquido_ultimomantenimiento, codigo_incidencias , sensores_cambiados_ultimomantenimiento, estado_baterias)
    # Actualizar el archivo destino con los datos leídos:

    # Actualizar la fecha del ultimo mantenimiento
    actualizar_fecha_destino(informe_destino, n_equipo, fecha_ultimomantenimiento, hoja_destino, hoja_destino_historico)

    # Actualizar el comentario del archivo destino con los datos leídos
    # actualizar_comentario_destino(informe_destino, n_equipo, comentario_ultimomantenimiento, hoja_destino, hoja_destino_historico)  
    
    # Actulizar el archivo destino con las nuevas medidas de metanol y liquido despues el mantenimiento
    # actualizar_metanol_destino(informe_destino, n_equipo, metanol_ultimomantenimiento, hoja_destino, hoja_destino_historico)
    # actualizar_liquido_destino(informe_destino, n_equipo, liquido_ultimomantenimiento, hoja_destino, hoja_destino_historico)

    # Actualizar el estado de los filtros de Lidar
    # actualizar_filtros_destino(informe_destino, n_equipo, codigo_incidencias, hoja_destino, hoja_destino_historico)

    #Actualizar el estado de la escobilla del Lidar
    # actualizar_escobilla_destino(informe_destino, n_equipo, codigo_incidencias, hoja_destino, hoja_destino_historico)

    # Actualizar las incidencias en el archivo destino
    # actualizar_incidencias_destino(informe_destino, n_equipo, codigo_incidencias, hoja_destino, hoja_destino_historico)

    # Actualizar el estado de las baterías en el archivo destino
    # actualizar_baterias_destino(informe_destino, n_equipo, codigo_incidencias, estado_baterias, hoja_destino, hoja_destino_historico)
    
    # Actualizar el estado de los sensores en el archivo destino
    # actualizar_sensores_destino(informe_destino, n_equipo, sensores_cambiados_ultimomantenimiento, hoja_destino, hoja_destino_historico)

    # Leer los correos del archivo de origen
    correos_destino, fecha_envio = leer_correos_origen(informe_mtto, hoja_mtto)
    # Enviar un correo de notificación
    # Escribimos el mensaje con el comentario del último mantenimiento, resaltando los errores a tener en cuneta para el próximo mantenimiento
    mensaje = f"Hola,\n\nSe ha realizado el mantenimiento del equipo {n_equipo} a dia {fecha_ultimomantenimiento}.\n\nComentario del último mantenimiento:\n{comentario_ultimomantenimiento}\n\n"
    if len(codigo_incidencias)>0:
        mensaje += "Por favor, ten en cuenta las siguientes incidencias para el próximo mantenimiento:\n"
        for incidencia in codigo_incidencias:
            mensaje += f"- {incidencia}\n"
    enviar_correo(
        destinatario=correos_destino,
        asunto=f"Información para el próximo mantenimientocdel equipo {n_equipo}",
        mensaje=mensaje, hora_envio=fecha_envio, cc=None, adjuntos=informe_mtto)



# Ejecutar la función principal
# Esta parte del código se ejecuta cuando el script se corre directamente
# Si se importa como un módulo, no se ejecutará automáticamente, ya que main() no se llamará.
if __name__ == "__main__":
    main()
    print("Datos actualizados correctamente.")
    # Este script actualiza el archivo destino con los datos del archivo de origen.
    # Asegúrate de que las rutas de los archivos y los nombres de las hojas sean correctos.