# Ejecución de Proyectos 2026 — Dashboard en vivo

Dashboard (`index.html`) que lee directamente de Google Sheets cada vez que se abre, y se
refresca solo cada 5 minutos. No hay que regenerar nada manualmente: cualquier cambio en la
hoja "Resumen" se refleja automáticamente.

## Importante: no abrir el archivo con doble clic

Google bloquea la carga de datos cuando la página se abre como archivo local (`file://`) —
es una restricción de seguridad (CORS) del lado de Google, no de este proyecto. El dashboard
**solo funciona si se sirve desde una URL real (http/https)**, por ejemplo GitHub Pages
(igual que el dashboard de Ventas de la carpeta principal).

Pasos sugeridos para publicarlo:

1. Sube esta carpeta (o todo el repositorio) a GitHub.
2. Activa GitHub Pages apuntando a la rama principal.
3. Comparte el link de GitHub Pages — cada vez que alguien lo abra, va a jalar los datos
   más recientes de la hoja de Google directamente.

Para probarlo localmente mientras tanto, sirve la carpeta con cualquier servidor estático
(no abrir el archivo directo), por ejemplo:

```
npx serve .
```

## Qué pestaña de Google Sheets usa

Por ahora solo la pestaña **"Resumen"** del archivo `EJECUCIÓN DE PROYECTOS_2026_KM.xlsx`.
El enlace publicado está en la constante `SHEETS.resumen.url` dentro de `index.html`. Cuando
se agregue otra pestaña al dashboard, hay que:

1. Publicarla en Google Sheets (Archivo → Compartir → Publicar en la web → elegir esa
   pestaña → formato CSV) para obtener su propio enlace con `gid`.
2. Agregarla al objeto `SHEETS` en `index.html` y crear su pestaña/sección correspondiente.

## Cómo se calculan las cifras

- **Se excluyen** del análisis financiero las filas de "NOTA DE CRÉDITO" y las filas sin
  cliente o sin datos (huecos en la numeración de factura).
- **Total facturado** se recalcula como `Base imponible + Exento + IVA + IGTF` por proceso,
  en lugar de usar la columna "TOTAL FACTURA" de la hoja — esa columna tiene errores de
  fórmula (`#REF!`) y algún desfase en al menos una fila de la hoja original. Recalcularlo
  desde las columnas base evita arrastrar esos errores al dashboard.
- **Efectividad (margen real)** = Utilidad real / (Base imponible + Exento). Esta fórmula se
  verificó contra el resumen manual de la propia hoja (40,87% general) y coincide.
- Las columnas **"Proyecto asignado a"** y **"Estatus"** se muestran tal cual estén en la
  hoja (hoy vacías) — a medida que se llenen en Google Sheets, aparecerán solas en la tabla
  de detalle sin tocar el código.

## Estructura

```
ejecucion-proyectos-2026/
└── index.html   — dashboard completo (HTML + CSS + JS), sin dependencias externas
                    salvo Google Fonts y el fetch al CSV publicado.
```
