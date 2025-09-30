import pandas as pd
import numpy as np

# Importar la librería numpy para manejar arrays y operaciones numéricas
# Importar la librería pandas para manejar DataFrames
# Importar la librería openpyxl para trabajar con archivos Excel



# Definimos nombre del archivo de origen y destino
informe_mtto = r'\\S31889004.services.dekra.com\DKI\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\LearningExcelLinks\informe_mtto.xlsx'
informe_destino = r'\\S31889004.services.dekra.com\DKI\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\LearningExcelLinks\ExcelPadre.xlsx'
# Definimos nombre de las hojas de origen y destino
hoja_mtto = 'Hoja1' #Nombre de la hoja del archivo de origen
hoja_destino = 'Lidar Windcube' #Nombre de la hoja del archivo destino
hoja_destino_historico = 'Historico' #Nombre de la hoja del archivo destino para el historico



# Importar las funciones de mantenimiento desde el módulo mantenimiento.py
# Asegúrate de que el archivo mantenimiento.py esté en el mismo directorio o en el PYTHONPATH
from mantenimiento import (leer_datos_origen, actualizar_fecha_destino, actualizar_comentario_destino, actualizar_metanol_destino, actualizar_liquido_destino)




def main():
    # Leer datos del archivo de origen:
    n_equipo, fecha_ultimomantenimiento, comentario_ultimomantenimiento, metanol_ultimomantenimiento, liquido_ultimomantenimiento = leer_datos_origen(informe_mtto, hoja_mtto)
    
    # Actualizar el archivo destino con los datos leídos:

    # Actualizar la fecha del ultimo mantenimiento
    actualizar_fecha_destino(informe_destino, n_equipo, fecha_ultimomantenimiento, hoja_destino, hoja_destino_historico)
    # Actualizar el comentario del archivo destino con los datos leídos
    # actualizar_comentario_destino(informe_destino, n_equipo, comentario_ultimomantenimiento, hoja_destino, hoja_destino_historico)  
    
    # Actulizar el archivo destino con las nuevas medidas de metanol y liquido despues el mantenimiento
    # actualizar_metanol_destino(informe_destino, n_equipo, metanol_ultimomantenimiento, hoja_destino, hoja_destino_historico)
    # actualizar_liquido_destino(informe_destino, n_equipo, liquido_ultimomantenimiento, hoja_destino, hoja_destino_historico)


# Ejecutar la función principal
# Esta parte del código se ejecuta cuando el script se corre directamente
# Si se importa como un módulo, no se ejecutará automáticamente, ya que main() no se llamará.
if __name__ == "__main__":
    main()
    print("Datos actualizados correctamente.")
    # Este script actualiza el archivo destino con los datos del archivo de origen.
    # Asegúrate de que las rutas de los archivos y los nombres de las hojas sean correctos.