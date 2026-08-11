# X4A — Formulario de Llaves (versión web)

Réplica en HTML/JS del Excel `Llaves_para_X4A.xlsx`, con las mismas listas
dependientes, validaciones de llave y cálculo de TRM — pero con los datos
guardados en un solo lugar compartido (Google Sheet), para que no se
pisen los cambios entre varias personas como pasaba en el Excel.

Incluye un **Dashboard** donde puedes ver las oportunidades cargadas,
seleccionar varias y enviarles a los responsables un correo confirmando
que su oportunidad ya quedó registrada en X4A (para evitar duplicados).

## Cómo está armado

```
x4a-app/
├── index.html          → la app (Formulario + Dashboard)
├── css/style.css
├── js/app.js            → toda la lógica (equivalente a las fórmulas del Excel)
├── js/config.js         → aquí pegas la URL de tu backend
├── data/export.json      → tabla "Export" del Excel (llaves válidas)
├── data/canales.json     → tabla "Canales" del Excel (8.048 registros, BCN)
└── backend/Codigo.gs     → backend en Google Apps Script (base de datos + correos)
```

## Paso 1 — Backend (Google Apps Script)

Esto reemplaza el Excel compartido: es donde quedan guardadas las
oportunidades, sin choques entre usuarios, y desde donde se disparan
los correos de notificación.

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja
   nueva. Llámala, por ejemplo, `X4A - Oportunidades`.
2. Menú **Extensiones → Apps Script**.
3. Borra todo el contenido del archivo `Código.gs` que se abre y pega
   ahí el contenido completo de `backend/Codigo.gs` de este proyecto.
4. Guarda (ícono de disco).
5. Arriba a la derecha, botón **Implementar → Nueva implementación**.
   - Tipo: **Aplicación web**.
   - Ejecutar como: **Yo (tu correo)**.
   - Quién tiene acceso: **Cualquier usuario** (para que tu equipo
     pueda usar el formulario sin iniciar sesión en Apps Script).
6. Autoriza los permisos que pida Google (es tu propio script).
7. Copia la **URL de la aplicación web** que te entrega al final
   (termina en `/exec`).

> La primera vez que alguien guarde una oportunidad, el script crea
> automáticamente la pestaña `Oportunidades` con los encabezados
> correctos — no necesitas prepararla a mano.

## Paso 2 — Conectar el frontend al backend

Abre `js/config.js` y reemplaza:

```js
const APPS_SCRIPT_URL = "PEGA_AQUI_TU_URL_DE_APPS_SCRIPT";
```

por la URL que copiaste en el paso anterior.

## Paso 3 — Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub (puede ser privado si tu plan
   lo permite, o público).
2. Sube todo el contenido de esta carpeta (`x4a-app/`) a la raíz del
   repositorio.
3. En el repositorio: **Settings → Pages**.
   - Source: `Deploy from a branch`
   - Branch: `main` / carpeta `/ (root)`
4. Espera 1-2 minutos; GitHub te da una URL tipo
   `https://tuusuario.github.io/tu-repo/`.
5. Comparte esa URL con tu equipo. Todos usan la misma página, y todos
   guardan en el mismo Google Sheet — sin sobrescribirse entre ellos.

## Notas sobre la réplica de la lógica del Excel

- **Canal → BCN**: se busca por coincidencia exacta (recortando
  espacios) contra `data/canales.json`, igual que el `XLOOKUP` del
  Excel contra la hoja `Canales`. Si hay más de un canal con el mismo
  nombre, se muestra el mismo aviso "OJO: hay N registros…".
- **Origen → Gobierno → Campaña**: listas dependientes tomadas de
  `data/export.json` (equivalente a la hoja `Export`).
- **TEMA final**: se arma igual que la fórmula original
  (`prefijo_Detalle_con_guiones_bajos`), con el mismo mensaje de error
  si falta el detalle.
- **Validación de llave** y **aviso de ambigüedad**: mismas reglas
  que las celdas `C15` y `C16` del Excel.
- **Ingreso estimado indexado**: `Valor cotizado × TRM del día ÷ TRM
  base`, igual que la celda `C22`.
- Los campos "Moneda" (US Dollar) y "País" (Colombia) quedan fijos,
  igual que en la plantilla original.

## Mantenimiento

- Si cambian las llaves válidas (Origen/Gobierno/Campaña/TEMA), edita
  `data/export.json` con el mismo formato y vuelve a subir el archivo.
- Si cambia el listado de canales/BCN, regenera `data/canales.json`
  desde la hoja `Canales` más reciente (columnas `Codigo Cliente BCN`
  y `Nombre Cliente`).
- Las opciones de "Situación Actual" están directamente en
  `js/app.js` (constante `SITUACIONES`) — edítalas ahí si cambian.
