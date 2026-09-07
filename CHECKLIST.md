# Checklist de desarrollo de PDFToolKit

## Fase 0 — Preparación y seguridad

- [x] Confirmar versión objetivo de Python y dependencias compatibles: Python 3.13.x de 64 bits.
- [ ] Crear/recuperar un repositorio Git y realizar un commit inicial. **Bloqueado: `.git` no permite escritura.**
- [x] Crear un entorno virtual reproducible con Python 3.13 en `.venv`.
- [x] Instalar y validar dependencias: `pip check`, compilación e imports principales correctos.
- [ ] Verificar que la aplicación arranca correctamente en Windows. **Pendiente de probar con una sesión gráfica.**
- [x] Añadir un `README` actualizado con instalación, uso y limitaciones reales.

## Fase 1 — Estabilización de la base

- [x] Añadir `if __name__ == "__main__":` y una función `main()`.
- [x] Separar la interfaz gráfica de la lógica de tratamiento de PDFs.
- [x] Sustituir APIs obsoletas de PyPDF2 (`PdfFileReader`/`PdfFileWriter`).
- [x] Eliminar imports y dependencias que no se utilizan en la base actual.
- [x] Centralizar la selección de archivos y la gestión de errores.
- [x] Evitar variables globales compartidas entre ventanas.
- [x] Crear directorios temporales seguros y borrar sus archivos al terminar.
- [x] Evitar que una operación pesada bloquee la interfaz mediante hilos de trabajo y callbacks en Tkinter.
- [x] Corregir textos, nombres y codificación de la interfaz.

## Fase 2 — Consolidar funcionalidades existentes

- [x] PDF a Word: cerrar correctamente el conversor incluso cuando hay errores.
- [x] Unir PDF: exigir al menos dos archivos y validar entradas/salida.
- [x] Dividir PDF: permitir elegir carpeta y patrón de nombres de salida.
- [x] PDF a PowerPoint: conservar proporciones y limpiar temporales.
- [x] Corregir la opción “Comprimir PDF” e implementar compresión real con MuPDF.
- [x] Eliminar la implementación duplicada de PDF a PowerPoint.
- [x] Añadir validación, estado de operación y mensajes de finalización para archivos grandes.

## Fase 3 — Funcionalidades prioritarias

- [ ] Comprimir PDF.
- [x] Rotar páginas.
- [x] PDF a JPG.
- [x] JPG a PDF.
- [x] Marca de agua.
- [x] Proteger PDF con contraseña.
- [x] Desbloquear PDF cuando sea técnicamente posible, usando la contraseña válida.
- [x] Ordenar páginas.
- [x] Enumerar páginas.

## Fase 4 — OCR y conversiones adicionales

- [x] Rehacer OCR usando PyMuPDF para renderizar páginas de forma fiable.
- [x] Documentar la instalación de Tesseract en Windows y sus idiomas.
- [x] Permitir seleccionar idioma de OCR.
- [x] PDF a Excel, extrayendo el texto por página en hojas independientes.
- [x] Word a PDF mediante Microsoft Office o LibreOffice.
- [x] Excel a PDF mediante Microsoft Office o LibreOffice.
- [x] PowerPoint a PDF mediante Microsoft Office o LibreOffice.
- [x] HTML a PDF mediante Microsoft Office o LibreOffice.
- [x] Detectar motores disponibles, permitir elegir cuando existen ambos y avisar cuando no existe ninguno.

## Fase 5 — Funciones avanzadas

- [x] Edición básica de PDF mediante metadatos.
- [ ] Firma digital real, diferenciándola de una imagen o marca visual.
- [x] Reparación de PDF mediante reconstrucción y limpieza de objetos.
- [ ] Conversión a PDF/A.
- [x] Comparación de dos PDFs mediante informe de diferencias de texto.
- [ ] Escaneo a PDF.

## Fase 6 — Calidad y distribución

- [x] Crear tests automatizados con PDFs de ejemplo.
- [ ] Probar PDFs vacíos, protegidos, grandes, escaneados y corruptos.
- [x] Añadir logging para diagnóstico.
- [x] Validar rutas de entrada, salida y sobrescritura accidental.
- [ ] Mejorar accesibilidad y navegación de la interfaz.
- [x] Generar ejecutable Windows reproducible mediante `build.ps1`.
- [ ] Probar el ejecutable en una instalación limpia.
- [ ] Actualizar README, capturas y lista de funcionalidades.

## Orden recomendado

1. Fase 0: preparar un entorno reproducible.
2. Fase 1: estabilizar y reorganizar el código.
3. Fase 2: dejar sólidas las funciones que ya existen.
4. Fase 3: añadir funcionalidades útiles de PDF.
5. Fases 4 y 5: conversiones y funciones avanzadas.
6. Fase 6: tests, documentación y ejecutable.
