# Migración de Macros VBA a Python - Automatización de Reportes

Este proyecto tiene como objetivo reemplazar tareas manuales en Excel (hechas con VBA) por scripts en Python que:

Actualizan automáticamente archivos Excel con conexiones de datos.  
Envían reportes automáticamente por correo con Outlook.

---

## Tecnologías utilizadas

- Python 3.x
- [xlwings](https://pypi.org/project/xlwings/) – para abrir y refrescar archivos Excel con macros
- [pywin32](https://pypi.org/project/pywin32/) – para enviar correos con Outlook

---

## Cómo usar

1. Clona este repositorio.
2. Coloca tu archivo `.xlsm` en la carpeta `/data/`.
3. Asegúrate de tener Outlook instalado y configurado.
4. Modifica el destinatario en `main.py` si es necesario.
5. Ejecuta el script:

```bash
python main.py
