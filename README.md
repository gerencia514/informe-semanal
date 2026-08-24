# Informe Semanal — Dashboard de Ventas

Dashboard estático (`index.html`) generado automáticamente a partir del archivo de Excel
de ventas. No necesita servidor ni instalar nada para verse — es una sola página web.

## Estructura del proyecto

```
INFORME SEMANAL/
├── index.html                                # El dashboard (generado, no editar a mano)
├── generar_reporte.ps1                       # Script que genera index.html desde el Excel
└── 2026.08.19_reporte de ventas.xlsx         # Tus datos reales (hoja "VENTAS 2026")
```

## Cómo actualizar el dashboard cada semana

1. Reemplaza el archivo de Excel en esta carpeta por la versión actualizada de la semana
   (mismo nombre, o cambia el nombre dentro de `generar_reporte.ps1` / pásalo como parámetro).
2. Abre una terminal en esta carpeta y ejecuta:

   ```
   powershell -ExecutionPolicy Bypass -File generar_reporte.ps1
   ```

3. Esto regenera `index.html` con los datos nuevos. Ábrelo haciendo doble clic para
   verlo, o súbelo a GitHub para que el link compartido se actualice (ver abajo).

Si tu archivo de Excel cambia de nombre cada semana, indícalo así:

```
powershell -ExecutionPolicy Bypass -File generar_reporte.ps1 -ExcelPath "nombre_del_archivo.xlsx"
```

## Qué calcula el script

- KPIs generales: ventas totales, base imponible + exento, IVA/IGTF, facturado y
  cobrado, facturado por cobrar, pendiente de facturar.
- Evolución mensual de ventas.
- Análisis por cliente (venta total, cobrado, por cobrar, sin facturar).
- Estado de facturación y cobranza.
- Antigüedad de cuentas por cobrar (0-30, 31-60, 61-90, +90 días, calculado respecto
  a la fecha en que corres el script).
- Top 10 operaciones por monto.

**Importante sobre el estado de cobranza:** la hoja de Excel trae unas filas de
resumen ya validadas ("FACTURADO Y COBRADO", "FACTURADO Y NO COBRADO",
"SIN FACTURAR NI COBRAR") que a veces no coinciden exactamente con lo que dice la
columna "COBRANZA" fila por fila (por ejemplo, cuando no se actualizó ese dato a
tiempo). El script usa esas filas de resumen como fuente de verdad para que el
dashboard siempre cuadre con lo que ya validas en Excel. Si esas filas de resumen no
existen o no cuadran, usa la columna COBRANZA como respaldo.

## Publicar el dashboard en GitHub Pages

Pendiente de completar cuando tengas tu usuario de GitHub — el objetivo es que el
link funcione desde cualquier equipo sin instalar nada. Pasos generales:

1. Crear un repositorio en GitHub (público).
2. Subir `index.html` (y opcionalmente el script y el Excel, si no te importa que
   sean públicos).
3. Activar GitHub Pages apuntando a la rama principal.
4. Cada vez que actualices el Excel y regeneres `index.html`, subir los cambios
   (`git add`, `git commit`, `git push`) para que el link público se actualice.

**Nota de privacidad:** al ser un repo público, cualquiera con el link podría ver
las cifras y nombres de clientes del reporte.
