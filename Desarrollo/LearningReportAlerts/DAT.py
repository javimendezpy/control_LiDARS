import win32com.client
import os
import csv
from datetime import datetime, timedelta
import pytz

# Configuración
CARPETA_DESTINO = r"C:\Users\A509132\Downloads\Sergi Gispert\adjuntos"
ZONA_HORARIA = pytz.timezone("Europe/Madrid")

def leer_y_agregar_dat(ruta):
    try:
        with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
            lineas = f.readlines()

        if len(lineas) < 5:
            print(f"El archivo {ruta} no tiene suficientes líneas para procesar.")
            return

        encabezado = next(csv.reader([lineas[0]]))
        if len(encabezado) < 2:
            print(f"Encabezado inválido en {ruta}")
            return

        id_estacion = encabezado[1].strip()
        carpeta_estacion = os.path.join(CARPETA_DESTINO, id_estacion)
        os.makedirs(carpeta_estacion, exist_ok=True) # Crear carpeta si no existe

        archivo_maestro = os.path.join(carpeta_estacion, f"{id_estacion}.csv")

        datos = [next(csv.reader([line])) for line in lineas[5:] if line.strip()]

        if not datos:
            print(f"No hay datos válidos en {ruta}")
            return

        with open(archivo_maestro, "a", newline='', encoding="utf-8") as f_out:
            writer = csv.writer(f_out)
            for fila in datos:
                writer.writerow([id_estacion] + fila)

        print(f"Datos añadidos a {archivo_maestro}")
    except Exception as e:
        print(f"Error al procesar el archivo {ruta}: {e}")


def guardar_adjuntos(email):
    print(f"Revisando adjuntos de: {email.Subject}")
    dat_encontrado = False

    for i in range(email.Attachments.Count):
        attachment = email.Attachments.Item(i + 1)
        filename = attachment.FileName
        if filename.lower().endswith(".dat"):
            dat_encontrado = True
            print(f"Encontrado archivo .dat: {filename}")

            # Guardar temporalmente en una ruta segura
            temp_path = os.path.join(CARPETA_DESTINO, f"temp_{filename}")
            attachment.SaveAsFile(temp_path)
            print(f"Guardado temporal: {temp_path}")

            try:
                with open(temp_path, "r", encoding="utf-8", errors="ignore") as f:
                    lineas = f.readlines()

                if len(lineas) < 5:
                    print(f"Archivo {filename} no tiene suficientes líneas.")
                    os.remove(temp_path)
                    continue

                encabezado = next(csv.reader([lineas[0]]))
                if len(encabezado) < 2:
                    print(f"Encabezado inválido en {filename}")
                    os.remove(temp_path)
                    continue

                id_estacion = encabezado[1].strip()
                carpeta_estacion = os.path.join(CARPETA_DESTINO, id_estacion)
                os.makedirs(carpeta_estacion, exist_ok=True)

                destino_final = os.path.join(carpeta_estacion, filename)
                os.replace(temp_path, destino_final)
                print(f"Movido a carpeta: {destino_final}")

                leer_y_agregar_dat(destino_final)

            except Exception as e:
                print(f"Error al mover o procesar {filename}: {e}")
                if os.path.exists(temp_path):
                    os.remove(temp_path)

    if not dat_encontrado:
        print("Este correo no contiene archivos .dat. No se procesa.")

def revisar_correos():
    print("Buscando correos de las últimas 24 horas...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    ahora = datetime.now().astimezone(ZONA_HORARIA)
    inicio_intervalo = ahora - timedelta(hours=24)

    for message in messages:
        recibido = message.ReceivedTime
        if isinstance(recibido, datetime):
            recibido_local = ZONA_HORARIA.localize(recibido.replace(tzinfo=None))
            if inicio_intervalo <= recibido_local <= ahora:
                if message.Attachments.Count > 0:
                    guardar_adjuntos(message)
                else:
                    print("Correo sin adjuntos. Ignorado.")
            elif recibido_local < inicio_intervalo:
                break
        else:
            print("Fecha de recepción no válida.")

try:
    revisar_correos()
except Exception as e:
    print(f"Error: {e}")
