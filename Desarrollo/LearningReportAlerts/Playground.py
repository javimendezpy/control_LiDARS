import win32com.client as win32
import os

def enviar_correo(destinatario, asunto, mensaje, cc=None, adjuntos=None):
    """
    Envía un correo usando Outlook instalado en Windows.
    - destinatario: str o lista de correos
    - asunto: str
    - mensaje: str (texto plano)
    - cc: str o lista de correos (opcional)
    - adjuntos: lista de rutas de archivos (opcional)
    """

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0 = MailItem

    # Destinatarios
    if isinstance(destinatario, list):
        mail.To = ";".join(destinatario)
    else:
        mail.To = destinatario

    # CC
    if cc:
        if isinstance(cc, list):
            mail.CC = ";".join(cc)
        else:
            mail.CC = cc

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
    mail.Send()
    print(f"✅ Correo enviado a {destinatarios_str}")


# Ejemplo de uso
if __name__ == "__main__":
    enviar_correo(
        destinatario=["jose.manuel.perez@dekra.com"],
        asunto="Prueba desde Python con win32com 3",
        mensaje="Hola, este correo fue enviado automáticamente desde Outlook usando Python.",
        # adjuntos=[r'\\S31889004.services.dekra.com\DKI\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\LearningReportAlerts\attachment1.xlsx'] # Asegúrate de que la ruta del archivo sea correcta
    )



