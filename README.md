# Actas de Bienes Públicos · CIIP

App web instalable (PWA) de la Gerencia de Bienes Públicos del Centro Internacional de Inversión Productiva. Sirve para emitir, corregir y controlar actas de asignación, desincorporación, movimiento y salida de bienes, con correlativos por tipo y año, revisiones y bitácora de cambios.

- **App:** https://carloskossik.github.io/ciip-actas-bienes/
- **Datos:** viven en un servidor privado de Google Apps Script con Google Sheets, que se instala aparte. Este repositorio solo contiene la interfaz: no guarda actas, usuarios ni claves.

## Instalar en un dispositivo

| Dispositivo | Cómo |
|---|---|
| Android (Chrome) | Abrir la dirección → menú ⋮ → **Instalar aplicación** o **Agregar a la pantalla principal** |
| iPhone / iPad (Safari) | Abrir la dirección → **Compartir** → **Agregar a pantalla de inicio** |
| Windows / Mac / Linux (Chrome o Edge) | Abrir la dirección → ícono **Instalar** en la barra de direcciones |

La app abre sin conexión y muestra los últimos datos sincronizados. Para emitir o corregir actas necesita internet, porque el servidor asigna el número.

## Archivos

- `index.html`: la app completa, sin dependencias externas.
- `config.js`: dirección del servidor (`apiUrl`, termina en `/exec`).
- `manifest.webmanifest`, `sw.js`, `icons/`: instalación y funcionamiento sin conexión.
