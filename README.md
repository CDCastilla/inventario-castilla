# Inventario · Clínica Dental Castilla

App móvil (Android, Chrome) para controlar el material de la clínica escaneando las cajas:
entradas, salidas (consumido, gabinete o caja quirúrgica), stock por lote y caducidad, avisos de mínimos
y reposición de gabinetes.

- **App**: página web instalable (PWA) publicada con GitHub Pages desde este repositorio.
- **Servidor**: Google Apps Script (`servidor/Codigo.gs`) pegado en la hoja de cálculo.
- **Datos**: una hoja de Google en la cuenta de la clínica. La hoja es también el informe y la copia.

No hay datos de pacientes en la app. Este repositorio es público: no contiene datos de la clínica ni claves
(la URL del servidor y la clave se escriben en el móvil, en Ajustes).

## Puesta en marcha (una vez)

### 1. Hoja de cálculo y servidor
1. Con la cuenta de Google de la clínica, crea una hoja de cálculo nueva llamada **Inventario Castilla**.
2. En la hoja: **Extensiones → Apps Script**. Borra lo que haya en `Código.gs` y pega el contenido de
   [`servidor/Codigo.gs`](servidor/Codigo.gs). Guarda (icono del disquete).
3. En Apps Script, **Configuración del proyecto** (rueda dentada) → zona horaria **(GMT+01:00) Madrid**.
4. Vuelve al editor, elige la función **`configurar`** en el desplegable y pulsa **Ejecutar**.
   Google pedirá permisos (hoja, Drive para las copias y conexión externa para exportar a Excel): acéptalos
   con la cuenta de la clínica. Si sale «Google no ha verificado esta aplicación», pulsa *Configuración
   avanzada → Ir a … (no seguro)*: es tu propio script.
5. Vuelve a la hoja y recárgala: aparece el menú **Inventario**. En **Inventario → Ver la clave de la app**
   tienes la clave.

### 2. Publicar el servidor
1. En Apps Script: **Implementar → Nueva implementación → tipo: Aplicación web**.
2. *Ejecutar como*: **Yo**. *Quién tiene acceso*: **Cualquier usuario**. Pulsa **Implementar**.
3. Copia la **URL de la aplicación web** (acaba en `/exec`).

### 3. Publicar la app (GitHub Pages)
1. En GitHub: **Settings → Pages → Build and deployment → Source: Deploy from a branch**,
   rama **main**, carpeta **/ (root)** → **Save**.
2. En unos minutos la app está en **https://cdcastilla.github.io/inventario-castilla/**.

### 4. En cada móvil Android
1. Abre la dirección de la app en **Chrome**.
2. Entra en **Ajustes** (en la pantalla de inicio de la app), pega la URL `/exec` y la clave, y pulsa
   **Guardar y probar**.
3. Menú de Chrome **⋮ → Instalar aplicación** (o *Añadir a pantalla de inicio*).
4. La primera vez que escanees, permite el uso de la cámara.

## Uso diario
- **Escanear** una caja abre su ficha (o pide darla de alta si el código es nuevo).
- **Entrada**: envases recibidos, lote y caducidad (se leen del DataMatrix), nº de factura o «pendiente».
- **Salida**: cantidad y destino. *Consumido al abrir* descuenta el envase entero; *por unidades* descuenta
  las que se usan. Lo que va a un gabinete sale del stock para el pedido y queda anotado en el gabinete.
- **Anular**: en la ficha del producto, junto a cada movimiento. Nada se borra.
- Sin cobertura se puede seguir trabajando: los movimientos esperan en el móvil y se envían solos.

## Actualizar
- **Servidor**: pega el `Codigo.gs` nuevo y en **Implementar → Gestionar implementaciones → editar (lápiz) →
  Versión: Nueva versión → Implementar**. La URL no cambia.
- **App**: los cambios se publican al subirlos a `main`. Cada versión sube `VERSION` en `sw.js` y el `?v=`
  de `index.html`; en el móvil puede hacer falta cerrar y abrir la app.

## Copias de seguridad
`configurar` programa una copia cada **lunes a las 7:00** en Drive (carpeta *Inventario Castilla - Copias*,
se guardan las 4 últimas). La app avisa si pasan más de 8 días sin copia. **Ajustes → Exportar a Excel**
descarga el inventario completo cuando se quiera.

## Hojas
| Hoja | Qué guarda |
|---|---|
| PRODUCTOS | Ficha de cada producto (categoría, armario, unidades por envase, modo de consumo, lote, mínimo) |
| CODIGOS | Código escaneado (GTIN/EAN) → producto |
| MOVIMIENTOS | Entradas, salidas, ajustes y anulaciones. Solo se añade |
| STOCK | Vista calculada por producto y lote (para consultar; se rehace sola) |
| DOTACION | Lo que debe tener cada gabinete (gabinete, producto, cantidad, días entre revisiones) |
| CONFIG | Categorías (modo y lote por defecto), armarios y destinos |
| FACTURAS | Facturas leídas (fase 2) |
| RECHAZADOS | Envíos que el servidor no aceptó, con el motivo |

## Pruebas
```
node --test pruebas/*.test.mjs
```
Lector de códigos GS1/EAN con etiquetas reales, cálculo de stock y un simulador del servidor.

## Tipografía
La app usa las fuentes del sistema con los colores de la marca. Las tipografías de la marca (Brockmann,
Recoleta) no se suben a este repositorio público por su licencia.
