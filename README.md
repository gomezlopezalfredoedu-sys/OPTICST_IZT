# OPTICST — V23 PWA

Base funcional: V23 de Boutique de Lentes Yael.

## Backend
Google Sheets:
1lpSxsrZp8w2QvtiU9pIQM5wyVKVznYw-ptyGsHZkhBOhRtcGhujip1j8

1. Crea/despliega este `Code.gs` en Apps Script.
2. Despliega como **Web app**.
3. Ejecutar como: **tú**.
4. Acceso: **cualquier persona que tenga el enlace** (según las opciones disponibles en tu cuenta).
5. Copia la URL que termina en `/exec`.
La URL `/exec` ya quedó configurada en `config.js`.

## GitHub Pages
Sube estos archivos al raíz del repositorio:

- index.html
- config.js
- manifest.json
- sw.js
- icon-192.png
- icon-512.png

## Instalación Android
Abrir la URL de GitHub Pages en Chrome. Cuando Chrome entregue el evento de instalación aparecerá el botón **Instalar OPTICST** dentro de la app.

No usar **“Instalar y crear acceso…”**, porque eso crea un acceso directo de Chrome y puede mostrar el icono de Chrome.

## Importante
El frontend conserva las llamadas `google.script.run` de V23 mediante un puente compatible; no se cambió la lógica de ventas, clientes, abonos, catálogos, tickets, PDF ni imágenes.
