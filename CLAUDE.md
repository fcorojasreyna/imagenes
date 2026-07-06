# Sistema CI - Servicios Vehiculares (Contexto para Claude Code)

## Qué es este sistema
Google Apps Script (GAS) web app desplegada como HtmlService. Es un SPA (Single Page Application) para gestión de taller automotriz, inventario de refacciones, consumos, proveedores, cotizaciones/órdenes de compra, herramientas, cronograma/agenda y dashboard multimodular.

## Arquitectura
- `Codigo.gs` — Backend GAS (~1584 líneas). Todas las funciones del servidor.
- `Main.html` — Shell SPA (~4146 líneas). Contiene TODO el JS funcional + CSS + estructura de navegación + lógica de todos los módulos.
- `Taller.html`, `Inventario.html`, `Consumos.html`, `Proveedores.html`, `Cotizador.html`, `Dashboard.html`, `Notificaciones.html` — Módulos HTML inyectados vía `incluir()` en el `#contenedor` del shell.

## Regla crítica de GAS/HtmlService
Los scripts dentro de HTML inyectado via `innerHTML` NO se ejecutan. Por eso TODO el JavaScript funcional vive en el `<script>` de `Main.html`. Los módulos `.html` solo contienen estructura HTML con atributos `onclick="window.funcionX()"`.

## Patrones de código usados
- Todas las funciones globales: `window.nombreFuncion = function() {...}` en Main.html
- Llamadas al backend: `google.script.run.withSuccessHandler(fn).withFailureHandler(fn).nombreFuncion(args)`
- Respuestas del backend: `{ exito: true/false, datos: [...], error/msj: "..." }`
- CSS variables: `--azul: #1E252B`, `--dorado: #B4A169`
- Font: Century Gothic
- UI: Bootstrap 5.3
- Formato de fechas: DD/MM/YYYY en todo el sistema

## Paleta de colores del Dashboard
```js
var _DP = ['#1E252B','#B4A169','#2d3a45','#c8ae7e','#3d5060','#8a7a4f','#6a8a7a','#c49a6a','#9a6a45','#4a6a5a'];
```
Siempre usar estos 10 tonos para charts. Nunca colores externos ni librerías de gráficas.

## Sistema de roles
Variable global `_rolActual` (string uppercase). Valores posibles:
- `'MECANICO'` — solo lectura cronograma, sin formularios de alta
- `'EJECUTIVO'` — operativo, puede registrar en todos los módulos
- `'GERENTE'` — puede autorizar, editar cronograma, ver panel pendientes
- `'SUBDIRECTORA'` — igual que GERENTE
- `'DIRECTORA'` — igual que GERENTE
- `'ADMIN'` — acceso total

```js
var rolesQueAutorizan = ['GERENTE','SUBDIRECTORA','DIRECTORA','ADMIN'];
```
Estos roles ven el panel de autorizaciones pendientes (`panelPendientesAuth`) y el historial de autorizaciones revisadas (`historialPendCard`) en Notificaciones.

## Variables globales de datos (definidas en Main.html ~línea 631)
```js
var _tallerDatos = [];       // datos del módulo Taller
var _inventarioDatos = [];   // datos del módulo Inventario
var _consumosDatos = [];     // datos del módulo Consumos
var _herramientasDatos = []; // datos del módulo Herramientas
var _dashCronDatos = [];     // datos de Cronograma para Dashboard (cargados al abrir Dashboard)
// _ocDatosGlobal — datos de OC, cargados por cargarDatosHistorialOC
// _provDatos — datos de Proveedores, cargados por cargarDatosProv
```

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

## Hojas de Google Sheets
Accesos, Vehiculos, Taller, Inventario, Consumos, Proveedores, OC, COMPARATIVA DOC, OC DOC, FORMATO ALTA VEHICULOS, Herramientas, Cronograma

## Funciones backend clave (Codigo.gs)
- `buscarInfoNuco(nuco)` — Busca vehículo por NUCO (col E=4), retorna todos los campos del vehículo
- `buscarInsumosPorTicket(ticket)` — Retorna refacciones DISPONIBLES de un ticket de inventario
- `obtenerHistorialNuco(nuco)` — Historial de taller y consumos por NUCO
- `registrarProveedor(d)` — Retorna `{ msj: "..." }`
- `eliminarFilaTaller(id)`, `eliminarFilaInventario(id)`, `eliminarFilaConsumo(id)`, `eliminarProveedor(id)` — Wrappers de `_eliminarRegistro(tipo, id)`
- `registrarOC(datosGenerales, cotizaciones, configCorreo, correo)` — Genera PDFs y guarda en Drive
- `obtenerDatosTaller()`, `obtenerDatosInventario()`, `obtenerDatosConsumos()`, `obtenerDatosProveedores()` — Lectura de hojas
- `obtenerDatosHerramientas()` — Lectura hoja Herramientas
- `obtenerDatosCronograma()` — Lectura hoja Cronograma (retorna `{ exito, datos }`)
- `obtenerDatosOC()` — Lectura hoja OC para historial

## Funciones frontend clave (Main.html) — referencia rápida por línea
| Función | Línea aprox. | Descripción |
|---------|-------------|-------------|
| `window.mostrarAlerta(msj)` | 407 | Modal de alerta genérico |
| `window.mostrarConfirmacion(msj, cb)` | 418 | Modal de confirmación con callback |
| `window.validarCampos(ids)` | 429 | Valida que campos no estén vacíos |
| `window.sanitizarNumero(v)` | 468 | Limpia string numérico → float |
| `window.formatearDatos(d)` | 474 | Formatea objeto de datos del vehículo |
| `window.procesarLogin()` | 486 | Login con hoja Accesos, setea `_rolActual` |
| `window.mostrarInicio(e)` | 535 | Carga panel de inicio con búsqueda rápida |
| `window._toggleNavMenu()` | 589 | Abre/cierra menú hamburger (móvil) |
| `window._cerrarNavMenu()` | 593 | Cierra menú hamburger |
| `window.cargarModulo(n, el)` | 598 | Carga módulo en #contenedor; incluye guard de cambios sin guardar |
| `window.cargarDatosTaller()` | 637 | Carga y renderiza tabla Taller |
| `window.cargarDatosInventario()` | 654 | Carga y renderiza tabla Inventario |
| `window.cargarDatosConsumos()` | 760 | Carga y renderiza tabla Consumos |
| `window.cargarDatosProv()` | 776 | Carga y renderiza tabla Proveedores |
| `window.ejecutarEliminar(modulo, id)` | 795 | Elimina registro de cualquier módulo |
| `window.verHistorial(nuco)` | 820 | Modal historial por NUCO |
| `window.exportarExcelTaller/Inventario/Consumos/Proveedores/Herramientas()` | 922+ | Exporta CSV por módulo |
| `window.cargarDatosHistorialOC()` | 970 | Carga historial de OC |
| `window.autocompletarNuco(valor, contexto, inputEl)` | 1050 | Autocompleta vehículo. Contextos: `'taller'` o `'oc'` |
| `window.abrirModalBuscadorProv()` | 1099 | Modal buscador de proveedores |
| `window.fijarProveedores()` | 1152 | Confirma lista de proveedores del cotizador |
| `window.soloAgencia()` | 1220 | Alias de `abrirModalCotizacionUnica` — cotización sin concurso |
| `window.procesarOC()` | 1557 | Genera OC completa con PDFs |
| `window._guardarBorradorOC()` | 1436 | Guarda borrador de OC en localStorage |
| `window._restaurarBorradorOC(datos)` | 1477 | Restaura borrador de OC |
| `window._marcarCambiosSinGuardar(formId)` | 1620 | Activa indicador de cambios sin guardar |
| `window.toggleVistaTaller(v)` | 1640 | Alterna tabla/formulario Taller |
| `window.calcularTotalTaller()` | 1667 | Recalcula servInterno + servExterno → total |
| `window.procesarTaller()` | 1676 | Guarda registro de taller |
| `window.toggleVistaInv(v)` | 1705 | Alterna tabla/formulario Inventario |
| `window.procesarInventario()` | 1727 | Guarda registro de inventario |
| `window.toggleVistaConsumos(v)` | 1743 | Alterna tabla/formulario Consumos |
| `window.procesarConsumo()` | 1773 | Guarda registro de consumo |
| `window.guardarProveedor()` | 1871 | Guarda proveedor nuevo/editado |
| `window.ordenarTabla(tbodyId, colIdx)` | 2018 | Ordenamiento de tablas por columna |
| `window.cargarDatosHerramientas()` | 2083 | Carga y renderiza tabla Herramientas |
| `window.guardarHerramienta()` | 2130 | Guarda herramienta nueva |
| `window.cargarCatalogos(intento)` | 2240 | Carga catálogos desde backend |
| `window.cargarDatosCronograma()` | 2398 | Carga y renderiza cronograma/agenda |
| `window.guardarCronograma()` | 2472 | Guarda registro de cronograma |
| `window.abrirEditCronograma(reg)` | 2535 | Abre modal edición cronograma |
| `window.cargarDatosNotificaciones(solobadge, avisoLogin)` | 2803 | Carga notificaciones y autorizaciones |
| `window.abrirModalAuth(id, tipo, nivel, ...)` | 3255 | Abre modal de autorización |
| `window.confirmarAutorizacion(decision)` | 3292 | Aprueba/rechaza autorización |
| `window.aplicarControlRol(rol)` | 3209 | Aplica visibilidad de elementos por rol |
| `window.iniciarDashboard()` | 3500 | Inicializa dashboard, carga todos los datos |
| `window._switchDashTab(tab)` | 3484 | Cambia tab del dashboard |
| `window._dashApplyFilter()` | 3474 | Aplica filtro mes/año al dashboard |
| `window._dashExportCSV()` | 3515 | Exporta datos del tab activo a CSV |
| `window._renderDashTaller()` | 3625 | Renderiza tab Taller del dashboard |
| `window._renderDashInventario()` | 3659 | Renderiza tab Inventario |
| `window._renderDashConsumos()` | 3688 | Renderiza tab Consumos |
| `window._renderDashProveedores()` | 3717 | Renderiza tab Proveedores |
| `window._renderDashOC()` | 3743 | Renderiza tab Órdenes de Compra |
| `window._renderDashHerramientas()` | 3769 | Renderiza tab Herramientas |
| `window._renderDashCronograma()` | 3800 | Renderiza tab Cronograma |
| `window._renderDashEjecutivos()` | 3831 | Renderiza tab Ejecutivos (cruza todos los módulos) |
| `window._renderDashMecanicos()` | 3901 | Renderiza tab Mecánicos (usa cronograma mecanico + mecanico2) |
| `window._precargarDatosGlobal()` | 3968 | Precarga todos los módulos al login para búsqueda global |
| `window._consultarInicio(modo)` | 3980 | Búsqueda en panel de inicio por ticket o NUCO |
| `window._busquedaGlobal(q)` | 4062 | Búsqueda global en tiempo real |
| `window.iniciarAutoRefreshNotif()` | 4131 | Inicia polling automático de notificaciones |

## Dashboard — particularidades importantes
- **9 tabs:** taller, inventario, consumos, proveedores, oc, herramientas, cronograma, ejecutivos, mecanicos
- **Filtro por mes/año:** selects `dashMesFiltro` y `dashAnoFiltro` — al cambiar llaman `_dashApplyFilter()`
- **`_dashFilter(data, fKey)`** — filtra array por mes y año del campo `fKey`. Acepta fechas DD/MM/YYYY.
- **costoTotal en Taller:** el campo `costoTotal` de la hoja viene como string formateado con "$" y comas. NUNCA parsearlo directamente. Siempre calcular como `servInterno + servExterno`.
- **Cronograma para Dashboard:** se carga en `_dashCronDatos` (no en el array de `cargarDatosCronograma`). Se precarga al abrir el Dashboard vía `obtenerDatosCronograma()`.
- **Ejecutivos:** cruza datos de taller (`quienRegistra`), OC (`quienRegistro`) y cronograma (`ejecutivo`).
- **Mecánicos:** usa campos `mecanico` y `mecanico2` del cronograma. Un registro puede contar para dos mecánicos.
- **SVG donuts:** implementados con `stroke-dasharray` + `transform="rotate(deg, 60, 60)"`. Sin librerías externas.
- **Exportar CSV:** genera `data:text/csv` con BOM UTF-8 (`﻿`), crea `<a>` temporal y lo clickea.

## Navbar responsive
- Hamburger (≡) visible solo en `≤768px` via CSS `.nav-hamburger { display:none }` + media query
- `navPillsContainer` se muestra/oculta con clase `open`
- `cargarModulo()` llama `_cerrarNavMenu()` al inicio para cerrar el menú en mobile
- `mostrarInicio()` también llama `_cerrarNavMenu()`

## Guard de cambios sin guardar
`cargarModulo()` verifica indicadores con IDs:
```
ind_sin_guardar_vistaFormTaller
ind_sin_guardar_vistaFormInventario
ind_sin_guardar_vistaFormConsumos
ind_sin_guardar_vistaFormProv
_indicadorBorradorOC
```
Si alguno está visible y con texto, muestra `confirm()` antes de navegar.

## Borrador OC (localStorage)
- Clave: `_draftOCKey()` → `'draftOC_' + usuario`
- Se guarda automáticamente con debounce de 800ms en cualquier cambio del formulario
- Al abrir el módulo OC se verifica si existe borrador y se ofrece restaurar

## Módulo Herramientas — campos del objeto
```
id, nombre, tipo, marca, serie, estado, estatus, lugarActual, responsable, fechaAdq, observaciones
```
`estatus`: Disponible / Prestada / En mantenimiento / Baja
`estado`: Bueno / Regular / Malo

## Módulo Cronograma — campos del objeto
```
id, ticket, nuco, marca, modelo, tipoTrabajo, mecanico, mecanico2, ejecutivo,
sede, estatusUnidad, fecha, horaEntrada, horaSalida, observaciones
```
- `estatusUnidad`: Ingresada / En proceso / Liberada / En espera de refacciones
- `tipoTrabajo`: catálogo de tipos de servicio
- Solo EJECUTIVO y GERENTE pueden editar; MECANICO solo lectura
- GERENTE ve panel adicional de estadísticas por mecánico

## Módulo Notificaciones / Autorizaciones
- Se recarga automáticamente con `iniciarAutoRefreshNotif()` (intervalo configurable)
- Autorizaciones tienen niveles: nivel 1 = Gerente, nivel 2 = Subdirectora/Directora
- `rolesQueAutorizan` = `['GERENTE','SUBDIRECTORA','DIRECTORA','ADMIN']`
- `historialPendCard` (autorizaciones revisadas) oculto para MECANICO y EJECUTIVO
- `panelPendientesAuth` oculto para MECANICO y EJECUTIVO
- Dupla de autorización: dos personas autorizan juntas (modal `toggleDuplaEdit`)

## Cotizador / Órdenes de Compra — flujo
1. Ejecutivo elige proveedores (modal buscador) → `fijarProveedores()`
2. Por proveedor llena precio unitario en matriz → `recalcularMatriz()`
3. `procesarOC()` → backend `registrarOC()` → genera PDFs en Drive
4. `soloAgencia()` = modo sin concurso, un solo proveedor (`abrirModalCotizacionUnica`)
5. Borrador se guarda automáticamente en localStorage

## CSS clases relevantes del Dashboard (Dashboard.html)
```css
.dash-kpi         /* tarjeta KPI: fondo blanco, borde izq dorado */
.dash-kpi .kpi-label  /* etiqueta uppercase pequeña */
.dash-kpi .kpi-value  /* valor grande azul */
.dash-kpi .kpi-sub    /* subtexto gris */
.dash-chart-card  /* contenedor de gráfica: fondo blanco, sombra */
.dash-chart-title /* título de gráfica: uppercase azul, borde inferior */
.dash-tab-btn     /* botón de tab: borde azul, activo = fondo azul */
.dash-filter-bar  /* barra de filtros: fondo blanco, flex */
.dash-export-btn  /* botón exportar: fondo dorado */
.dash-tabla       /* tabla interna del dashboard */
```

## Particularidades y bugs conocidos
- **costoTotal Taller:** campo crudo con formato `$XX,XXX.XX` — usar `servInterno + servExterno` siempre
- **Scripts en innerHTML:** nunca funcionan; todo JS va en `<script>` de Main.html
- **autocompletarNuco contextos:** `'taller'` rellena campos con prefijo `taller_*`; `'oc'` rellena campos con prefijo `oc_*`
- **Cronograma fecha:** campo `fecha` en formato DD/MM/YYYY; `_dashParseDate()` lo convierte a Date
- **Exportar Excel:** en realidad es CSV con extensión `.csv`; se abre bien en Excel
- **Bootstrap en GAS:** se carga desde CDN — si hay problemas de red, el UI puede verse sin estilos

## Fixes aplicados (historial)
1. Creada función `buscarInfoNuco` (no existía en backend)
2. Renombrada `buscarInsumoTicket` → `buscarInsumosPorTicket` + formato retorno corregido
3. Renombrada `obtenerHistorialVehiculo` → `obtenerHistorialNuco`
4. Agregadas funciones delete individuales (`eliminarFilaTaller`, etc.)
5. `registrarProveedor` ahora retorna `{ msj: "..." }` en lugar de string plano
6. Corregidos IDs de campos en `autocompletarNuco` para contextos 'taller' y 'oc'
7. Agregada función `soloAgencia()` (alias de `abrirModalCotizacionUnica`)
8. Navbar responsive con hamburger para ≤768px
9. Guard de cambios sin guardar en `cargarModulo()`
10. `historialPendCard` y `panelPendientesAuth` ocultos para roles no autorizantes
11. Dashboard completo 9 tabs: taller, inventario, consumos, proveedores, OC, herramientas, cronograma, ejecutivos, mecánicos
12. Filtro por mes/año en dashboard con `_dashFilter()`
13. Exportar CSV desde dashboard (`_dashExportCSV()`)
14. Tab Ejecutivos: cruza todos los módulos para tracking por ejecutivo
15. Tab Mecánicos: tracking de rendimiento por mecánico desde cronograma
16. Bug costoTotal: ahora se calcula como `servInterno + servExterno` (no del campo crudo)
17. Búsqueda global en panel de inicio por ticket y NUCO

## Git / Despliegue
- Rama activa: `claude/repository-file-access-piuxo7`
- Push requiere PAT con scope `repo` configurado como remote URL con credenciales
- Comando: `git remote set-url origin https://fcorojasreyna:<TOKEN>@github.com/fcorojasreyna/imagenes.git`
- Repositorio: `github.com/fcorojasreyna/imagenes`
