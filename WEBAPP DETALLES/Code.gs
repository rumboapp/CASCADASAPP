/**
 * ============================================================================
 * CASCADAS HOTEL — DETALLES CASCADAS
 * ============================================================================
 * Generador de tarjetas de atencion (aniversario, cumpleanos, bienvenida,
 * estadia de regalo, gift card de noches, welcome drink, disculpas, etc.)
 * y cupones canjeables. Cada tarjeta se arma en el navegador y se descarga
 * como PDF listo para imprimir en formato diptico pequeno ("French fold"):
 * se imprime por UNA sola cara, se dobla dos veces y queda una tarjeta A6
 * que se abre como tarjeta de saludo. Al ser una sola cara no hay problema
 * de alineacion frente/reverso.
 *
 * No usa planilla: no guarda datos, todo se genera al momento. Por eso el
 * unico punto de entrada es doGet (sirve el Index) — no hay backend de datos.
 * ============================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Detalles Cascadas — Cascadas Hotel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
