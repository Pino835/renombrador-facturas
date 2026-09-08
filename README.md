# Renombrador de Facturas

Automatización de escritorio (Python + OCR) para el flujo "factura escaneada → archivo renombrado" en procesos de cuentas por cobrar.

## Qué hace

1. Revisa uno por uno los PDF escaneados de una carpeta de entrada.
2. Extrae por OCR (Tesseract) el código de cliente y el número de factura de cada documento.
3. Si logra leer ambos datos con suficiente confianza, guarda una copia renombrada (`[codigo_cliente]-[factura].pdf`) en una carpeta de resultados organizada por fecha.
4. Si no logra leerlos con confianza, abre una ventana con los recortes del documento para completarlos a mano.
5. Registra cada resultado (automático o manual) en un log.
6. Usa una lista de códigos de cliente ya confirmados para autocorregir confusiones típicas del OCR (como "O" vs "0"), sin necesidad de revisión manual repetida.

## Tecnologías

- Python
- PyMuPDF (lectura de PDF)
- Tesseract OCR (vía `pytesseract`)
- Pillow
- Tkinter (interfaz de revisión manual)

## Cómo correr

1. `ejecutar_1_instalar_dependencias.bat` — una sola vez (o al cambiar de equipo). Instala Python y Tesseract-OCR si faltan, todo a nivel de usuario (sin permisos de administrador).
2. `ejecutar_2_renombrador.bat` — procesa los documentos de la carpeta de entrada.

Más detalle en `LEEME_PRIMERO.txt` y `COMO_FUNCIONA.txt`.

## Nota

Este proyecto se desarrolló como parte de un rol de automatización de procesos administrativos, dirigiendo herramientas de IA (Claude, Claude Code) para el diseño, implementación y documentación de la solución. El código aquí publicado no incluye datos, configuraciones ni documentos reales de ninguna organización.
