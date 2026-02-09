# Log de cambios

## 0.1.0
- Añadida la app local en HTML, CSS y JavaScript para importar CSV, filtrar por NIF y generar etiquetas Apli 01213.
- Incluida la descarga de un documento Word compatible y vista previa de etiquetas.

## 0.1.1
- Actualizado el formato de etiquetas a Apli 01273 (70 × 37 mm, 24 por hoja) en textos, vista previa y documento Word.
- Cambiada la versión visible de la app y el nombre del archivo de descarga.

## 0.2.0
- Rediseñado el flujo para cargar un único CSV de alumnos, visualizar los datos y abrir un modal para el cruce con un segundo CSV.
- Añadido el modal de coincidencias con NIF en verde, resumen de coincidencias y descarga de Word filtrado.
- Actualizados estilos, tablas y versión visible de la app.

## 0.3.0
- Añadida la importación de archivos XLSX con la misma estructura que los CSV, manteniendo las validaciones de columnas.
- Actualizados los textos, versionado visible y formatos aceptados para CSV/XLSX.

## 0.4.0
- Ajustada la generación de Word a una tabla APLI 01273 con “Alumno” y “Dirección completa” por celda.
- Añadida la descarga de archivo para combinar correspondencia APLI 01273 desde el modal.
- Actualizada la versión visible de la app y la previsualización de etiquetas.

## 0.4.1
- Corregida la descarga para generar un archivo Word compatible con HTML en formato .doc.
- Respetados los saltos de línea en la vista previa y en el documento Word para mantener el formato de etiquetas.
- Actualizada la versión visible de la app.

## 0.4.2
- Ajustado el HTML generado para Word con márgenes, tamaños y centrado alineados al formato del DOCX de referencia.
- Añadidos párrafos de línea en etiquetas con separación y negrita del nombre para mantener la alineación.
- Actualizada la versión visible de la app.

## 0.4.3
- Ajustado el HTML de etiquetas para respetar márgenes, alto de fila, ancho de columnas y sangrías del DOCX de referencia.
- Añadido botón para exportar a PDF con la misma maqueta de impresión que el Word.
- Actualizada la versión visible de la app.

## 0.4.4
- Igualadas las medidas APLI 01273 (márgenes, tabla y celdas) con el HTML de referencia para Word/PDF.
- Eliminados los anchos irregulares, el colgroup y el margin-left negativo en la tabla de etiquetas.
- Actualizada la versión visible de la app.
