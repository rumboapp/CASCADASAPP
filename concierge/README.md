# Cascadas Concierge

Centro operacional del hotel construido sobre **Google Apps Script + Google Sheets**.
No es una agenda ni una app de reservas tradicional: reemplaza llamadas, radios y
coordinación verbal entre huéspedes y personal.

## Archivos

| Archivo | Rol |
|---------|-----|
| `Setup.gs` | **Acto 1** — Crea el Spreadsheet con las 14 hojas, columnas y datos iniciales (incluye la carta completa de 132 productos). |
| `Code.gs` | **Acto 2** — Backend: APIs, motor de reservas, prepedidos, CRUD de carta/categorías, historial, seguridad y validaciones server-side. |
| `Index.html` | **Acto 2** — Frontend SPA autocontenido (CSS + HTML + JS inline) con vistas de Huésped (QR) y Staff. |
| `appsscript.json` | Manifiesto del proyecto Apps Script (webapp anónima, zona horaria de Chile). |

## Puesta en marcha

1. Crea un proyecto nuevo en [script.google.com](https://script.google.com).
2. Copia el contenido de `Setup.gs`, `Code.gs` y `Index.html` (este último como archivo HTML llamado `Index`).
3. Reemplaza el manifiesto con `appsscript.json` (activa "Mostrar appsscript.json" en Configuración del proyecto).
4. Ejecuta la función **`crearBaseDeDatos()`** una sola vez. Autoriza los permisos.
   - Se crea el Spreadsheet y su ID queda guardado en `ScriptProperties` (`SPREADSHEET_ID`), por lo que `Code.gs` lo usa automáticamente.
5. **Implementar → Nueva implementación → Aplicación web** (ejecutar como tú, acceso "Cualquiera").
6. Comparte la URL de la webapp.

## Accesos

- **Huésped**: se accede vía QR con `?hab=XXX` (query string). Ej: `.../exec?hab=105`.
  Si no hay `hab`, el huésped puede ingresar el número manualmente.
- **Staff**: botón "Soy Personal" → email registrado en la hoja `Usuarios`.
  Usuarios semilla: `admin@cascadas.cl` (ADMINISTRADOR), `recepcion@cascadas.cl` (RECEPCION), `cocina@cascadas.cl` (COCINA).

## Reglas de negocio clave

- **Motor único de reservas**: todos los servicios usan `crearReserva`; las diferencias (horario, capacidad, duración, aprobación) se parametrizan en la hoja `Servicios`.
- **Tinaja / Bicicletas / Masajes**: nacen en estado `Pendiente aprobación`.
- **Prepedido**: modificable hasta `PREPEDIDO_MINUTOS_LIMITE` (30 min) antes de la reserva. Sólo accesible desde un link discreto en "Mis Reservas".
- **Desayuno Habitación**: continental, no es a pedido plato por plato; admite notas libres; habilitable desde Configuración.
- **Soft delete**: productos y categorías nunca se borran físicamente (se marca `Visible`/`Disponible` en FALSE) para conservar el historial.
- **Historial de Pedidos**: sólo consulta interna (ADMIN/RECEPCIÓN). No genera comprobantes; cada búsqueda se audita en `HistorialPedidos`.
- **Sin fotos de platos**: la carta muestra nombre, descripción, precio y etiquetas. No existe columna `ImagenURL`.

Todo horario, capacidad y precio se lee de `Configuracion` / `Servicios`: nada está hardcodeado.
