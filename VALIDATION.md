# Hosted release / Versión alojada — 1.2.0

- `npm run check`: TypeScript, lint, 10 tests, production build, manifest validation, and packaging passed.
- `npm run build:hosted`: passed; hosted URLs, domain, icon files, manifest validation, ZIP manifest, and English translation verified.
- GitHub Pages deployment workflow publishes only `dist/` after checks pass.
- The local Office installation is not modified by building or publishing. A hosted package must be installed once to switch an existing local registration.
- Visual verification inside Word after installing the hosted package remains a manual check.

Las comprobaciones de compilación y paquete alojado pasaron. Publicar no cambia el complemento instalado. Para migrar una instalación local es necesario instalar el paquete alojado una vez. La comprobación visual en Word queda pendiente para esa instalación.

---

# Validation / Validación — 1.1.0

Validated on Windows with Node.js 26.11.1 and npm 11.20.0. GitHub Actions runs the project checks on Windows with Node.js 22.

Verificado en Windows con Node.js 26.11.1 y npm 11.20.0. GitHub Actions ejecuta las comprobaciones en Windows con Node.js 22.

- Clean dependency installation / instalación limpia: `npm ci` passed.
- TypeScript and lint / TypeScript y estilo: passed.
- Automated tests / pruebas automáticas: 10 passed, including English pane rendering, insertion feedback, Spanish fallback, manifest translation paths, code preservation, formatting, and Word insertion behavior.
- Production build / compilación optimizada: passed.
- Office manifest validation / validación del manifiesto: passed.
- English localization: validated against Microsoft's localization schema; included in the app package and compiled output.
- Browser authentication CLI: read-only authentication check passed with Agents Toolkit CLI 1.1.17; account details are excluded from this repository.

The installed Spanish version was preserved by preparing this release in a separate directory. Its existing HTTPS server continued to serve version 1.0.2 without additional languages.

La versión instalada en español se conservó preparando esta publicación en una carpeta separada. Su servidor HTTPS continuó sirviendo la versión 1.0.2 sin idiomas adicionales.

The English pane was tested with a simulated Office display language. A visual check of the localized ribbon in an English Word installation remains a manual verification; this release was not installed over the existing Spanish installation.

El panel en inglés se probó simulando el idioma de Office. La comprobación visual de la cinta traducida en Word en inglés queda como verificación manual; esta versión no se instaló sobre la instalación existente en español.
