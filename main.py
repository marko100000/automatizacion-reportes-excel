import xlwings as xw
import win32com.client as win32
import os
from datetime import date
import time

# Ruta del archivo
ruta_archivo = "data/reporte_ventas.xlsm"
# Destinatario
destino = "MARKO.NAVEDA@UPSJB.EDU.PE"

def actualizar_excel(path):
    print("🔄 Abriendo Excel y actualizando conexiones...")
    wb = xw.Book(path)
    wb.api.RefreshAll()
    time.sleep(5)  # Esperar que se actualicen las conexiones
    wb.save()
    wb.close()
    print("Excel actualizado y guardado.")

def enviar_correo(destinatario, archivo):
    print("📤 Preparando correo...")
    outlook = win32.Dispatch('outlook.application')
    correo = outlook.CreateItem(0)
    
    correo.To = destinatario
    correo.Subject = f"Reporte de Ventas - {date.today().strftime('%d-%m-%Y')}"
    correo.Body = "Estimado, adjunto el reporte de ventas. Quedo atento a tus comentarios."
    correo.Attachments.Add(os.path.abspath(archivo))
    correo.Send()
    
    print("Correo enviado exitosamente.")

# Ejecutar todo
if __name__ == "__main__":
    print("Iniciando automatización de reporte...\n")
    actualizar_excel(ruta_archivo)
    enviar_correo(destino, ruta_archivo)
    print("\n Proceso completado con éxito.")