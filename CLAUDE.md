# Sistema CI - Servicios Vehiculares (Contexto para Claude Code)

## Qué es este sistema
Google Apps Script (GAS) web app desplegada como HtmlService. Es un SPA (Single Page Application) para gestión de taller automotriz, inventario de refacciones, consumos, proveedores y cotizaciones/órdenes de compra.

## Arquitectura
- `Codigo.gs` — Backend GAS (862 líneas). Todas las funciones del servidor.
- `Main.html` — Shell SPA (1224 líneas). Contiene TODO el JS funcional + CSS + estructura de navegación.
- `Taller.html`, `Inventario.html`, `Consumos.html`, `Proveedores.html`, `Cotizador.html`, `Dashboard.html` — Módulos HTML inyectados vía `incluir()` en el `#contenedor` del shell.

## Regla crítica de GAS/HtmlService
Los scripts dentro de HTML inyectado via `innerHTML` NO se ejecutan. Por eso TODO el JavaScript funcional vive en el `<script>` de `Main.html`. Los módulos `.html` solo contienen estructura HTML con atributos `onclick="window.funcionX()"`.

## Patrones de código usados
- Todas las funciones globales: `window.nombreFuncion = function() {...}` en Main.html
- Llamadas al backend: `google.script.run.withSuccessHandler(fn).withFailureHandler(fn).nombreFuncion(args)`
- Respuestas del backend: `{ exito: true/false, datos: [...], error/msj: "..." }`
- CSS variables: `--azul: #1E252B`, `--dorado: #B4A169`
- Font: Century Gothic
- UI: Bootstrap 5.3

## Sheet "Vehiculos" — índices de columna (0-based)
| Campo | Columna | Índice |
|-------|---------|--------|
| Serie | C | 2 |
| NUCO | E | 4 |
| Responsable | F | 5 |
| Departamento | I | 8 |
| Marca | T | 19 |
| Clase | U | 20 |
| Línea | V | 21 |
| Modelo | W | 22 |
| Color | X | 23 |
| Placa | Y | 24 |
| Razón Social | AC | 28 |
| Sede | AH | 33 |
| Oficina | AI | 34 |

## Funciones backend clave (Codigo.gs)
- `buscarInfoNuco(nuco)` — Busca vehículo por NUCO (col E=4), retorna todos los campos del vehículo
- `buscarInsumosPorTicket(ticket)` — Retorna refacciones DISPONIBLES de un ticket de inventario
- `obtenerHistorialNuco(nuco)` — Historial de taller y consumos por NUCO
- `registrarProveedor(d)` — Retorna `{ msj: "..." }`
- `eliminarFilaTaller(id)`, `eliminarFilaInventario(id)`, `eliminarFilaConsumo(id)`, `eliminarProveedor(id)` — Wrappers de `_eliminarRegistro(tipo, id)`
- `registrarOC(datosGenerales, cotizaciones, configCorreo, correo)` — Genera PDFs y guarda en Drive
- `obtenerDatosTaller()`, `obtenerDatosInventario()`, `obtenerDatosConsumos()`, `obtenerDatosProveedores()` — Lectura de hojas

## Funciones frontend clave (Main.html)
- `window.autocompletarNuco(valor, contexto)` — Autocompleta campos del vehículo. Contextos: `'taller'` o `'oc'`
- `window.soloAgencia()` — Modo cotización única sin concurso de proveedores
- `window.toggleVistaTaller/Inventario/Consumos/Prov/OC(vista)` — Alterna entre tabla y formulario
- `window.procesarOC()` — Procesa y genera la orden de compra
- `window.fijarProveedores()` — Confirma lista de proveedores del cotizador
- `window.ordenarTabla(tbodyId, colIndex)` — Ordenamiento de tablas

## Hojas de Google Sheets
Accesos, Vehiculos, Taller, Inventario, Consumos, Proveedores, OC, COMPARATIVA DOC, OC DOC, FORMATO ALTA VEHICULOS

## Fixes aplicados (historial)
1. Creada función `buscarInfoNuco` (no existía en backend)
2. Renombrada `buscarInsumoTicket` → `buscarInsumosPorTicket` + formato retorno corregido
3. Renombrada `obtenerHistorialVehiculo` → `obtenerHistorialNuco`
4. Agregadas funciones delete individuales (`eliminarFilaTaller`, etc.)
5. `registrarProveedor` ahora retorna `{ msj: "..." }` en lugar de string plano
6. Corregidos IDs de campos en `autocompletarNuco` para contextos 'taller' y 'oc'
7. Agregada función `soloAgencia()` que faltaba en Main.html
