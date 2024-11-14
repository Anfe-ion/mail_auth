import json
import win32com.client as win32
import logging

# Configurar el logger
logging.basicConfig(
    filename="./logs/correo_envios.log",  # Archivo donde se guardarán los logs
    level=logging.INFO,            # Nivel de log (INFO para registrar envíos)
    format="%(asctime)s - %(levelname)s - %(message)s",  # Formato del mensaje de log
    datefmt="%Y-%m-%d %H:%M:%S"    # Formato de fecha y hora
)

try:
    # Cargar datos del archivo JSON
    with open("data/datos.json", "r", encoding="utf-8") as file:
        data = json.load(file)

    # Leer el contenido del archivo HTML
    with open("index.html", "r", encoding="utf-8") as file:
        base_html_content = file.read()

    # Configuración de Outlook
    outlook = win32.Dispatch("Outlook.Application")

    # Enviar el correo a cada destinatario en el JSON
    for recipient in data["recipients"]:
        # Crear las filas de la tabla en HTML usando los datos de 'tabla' de cada destinatario
        filas_tablas = ""
        for item in recipient["tabla"]:
            filas_tablas += f"""
            <tr>
                <td style="border: 1px solid black; padding: 8px;">{item['codigo_sap']}</td>
                <td style="border: 1px solid black; padding: 8px;">{item['codigo']}</td>
                <td style="border: 1px solid black; padding: 8px;">{item['description']}</td>
                <td style="border: 1px solid black; padding: 8px;">{item['unidades_fac']}</td>
                <td style="border: 1px solid black; padding: 8px;">{item['costo_total_fact']}</td>
                <td style="border: 1px solid black; padding: 8px;">{item['unidades_falt']}</td>
            </tr>
            """

        # Reemplazar placeholders en el contenido HTML con los valores del destinatario actual
        html_content = base_html_content.format(
            name=recipient["name"],
            codi_camp=recipient["codi_camp"],
            por_pen=recipient["por_pen"],
            ticket=recipient["ticket"],
            filas_tablas=filas_tablas
        )

        try:
            # Crear y enviar el correo
            mail = outlook.CreateItem(0)
            mail.To = recipient["email"]
            mail.Subject = f"Reporte de Cierre de {recipient['codi_camp']}"
            mail.HTMLBody = html_content
            mail.Send()
            
            # Registrar el envío exitoso en el archivo de log
            logging.info(f"Correo enviado con éxito a {recipient['name']} ({recipient['email']}).")

        except Exception as e:
            # Registrar errores específicos de envío en el archivo de log
            logging.error(f"Error al enviar correo a {recipient['name']} ({recipient['email']}): {e}")

except Exception as e:
    logging.critical(f"Error crítico al procesar los datos o enviar correos: {e}")