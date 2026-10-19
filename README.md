# WordCodeAddin

**Español** | [English (United States)](README.en.md)

Bloques de código editables con resaltado de sintaxis para Microsoft Word. La interfaz sigue el idioma de Word: español formal o inglés de Estados Unidos; otros idiomas usan español.

**[Descargar e instalar el complemento](https://reyhatsat.github.io/WordCodeAddin/)**

## Instalar la versión alojada

Esta versión carga el panel desde GitHub Pages. **No necesita Node.js, Start.cmd ni un servidor local para utilizarla.** Necesita conexión a Internet, Word compatible con manifiestos unificados y una cuenta de Microsoft 365 cuya organización permita aplicaciones personalizadas. El manifiesto conserva `devPreview` del proyecto original; la compatibilidad depende de su versión y canal de Office.

1. Descargue [wordcodeaddin.zip](https://reyhatsat.github.io/WordCodeAddin/wordcodeaddin.zip). **No lo extraiga.**
2. Guarde sus documentos y cierre Word.
3. Abra Microsoft Teams con la misma cuenta que utiliza en Word. Seleccione **Aplicaciones → Administrar sus aplicaciones → Cargar una aplicación → Cargar una aplicación personalizada**.
4. Seleccione el ZIP descargado, pulse **Agregar** y cierre Teams. No necesita abrir el complemento en Teams.
5. Abra Word y un documento editable. Seleccione **Insertar → Código → Insertar código**. Si no aparece, busque WordCodeAddin en **Inicio → Complementos**. Puede tardar unos minutos en aparecer.

Si no dispone de la opción para cargar aplicaciones, su administrador deberá permitirla o distribuir el paquete mediante los canales de Microsoft 365 de su organización. Consulte el [procedimiento oficial de Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest#sideload-through-the-teams-app-store) y los [requisitos de compatibilidad](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview).

**Después de instalarlo, solo tiene que abrir Word.** El panel y los iconos se descargan de GitHub Pages cuando Word los necesita. La dirección de Internet debe seguir disponible.

### Si ya tiene la versión local

La publicación de esta web no modifica su instalación. Para pasar a la versión alojada, debe instalar el ZIP anterior una vez; ambas versiones utilizan el mismo identificador de complemento. No necesita iniciar el servidor local para la versión alojada. Si Word conserva el registro anterior, cierre sus ventanas y revise el procedimiento de actualización de caché de Microsoft antes de volver a instalar. No borre documentos ni carpetas completas de Office para actualizarlo.

## Uso

Pegue su código, seleccione un lenguaje y el tema Claro u Oscuro, sitúe el cursor en Word y pulse **Insertar en Word**. El bloque se añade al final de la selección sin reemplazar el texto existente y permanece editable.

Incluye C#, Python, JavaScript, TypeScript, Java, HTML, CSS, SQL, JSON, Bash y Texto plano, con fuente Consolas de 10 puntos, colores, fondo, bordes, espacios, saltos de línea y tabulaciones. El complemento no traduce ni ejecuta el código; el procesamiento ocurre en su equipo. Vuelva a abrir el panel después de cambiar el idioma de la interfaz de Word.

## Solución de problemas

- **No aparece el botón:** confirme que Word admite el manifiesto unificado, que inició sesión con la misma cuenta y que la organización permite aplicaciones personalizadas. Busque WordCodeAddin en Inicio → Complementos y reinicie Word.
- **No aparece Cargar una aplicación personalizada:** solicite ayuda al administrador de Microsoft 365. Descargar otra vez el proyecto no cambia las políticas de la organización.
- **Panel sin conexión:** compruebe su acceso a Internet y a `reyhatsat.github.io`. La versión alojada no utiliza localhost ni certificados de desarrollo.
- **Todavía intenta abrir localhost:** Word conserva la instalación local; instale el paquete alojado y reinicie Word. Cambiar el repositorio en GitHub no actualiza por sí solo los registros de Word.
- **Documento protegido:** permita la edición o utilice un documento editable.

## Desarrollo local (opcional)

Esta sección es para modificar el código. Requiere [Node.js 22.15 o posterior](https://nodejs.org/es/download), Word para Windows y permisos para instalar complementos personalizados.

Descargue el código con **Code → Download ZIP**, extráigalo y haga doble clic en **Install.cmd**. Alternativamente, ejecute estas órdenes en la carpeta que contiene `package.json`:

```powershell
npm ci
npm run signin:browser
npm run start:desktop:word:browser
```

Use la misma cuenta que utiliza en Word. Acepte el certificado HTTPS local si se solicita. Para los siguientes usos de la versión de desarrollo, abra **Start.cmd** o ejecute `npm run dev-server` y mantenga el servidor funcionando. Estos scripts instalan la variante local, no la alojada.

El inicio por navegador incluye una alternativa al error nativo `platform_broker_error` de Windows, verificada con Agents Toolkit CLI 1.1.17. Solo modifica sus procesos hijos; no altera Windows ni las dependencias. Si se actualiza el Toolkit, deberá volver a comprobarse. Para errores de certificado local, ejecute `npx office-addin-dev-certs install`. Use el inicio normal sin depuración para evitar la pausa de WebView.

## Compilar y publicar

```powershell
npm run check
npm run build:hosted
```

La primera orden comprueba TypeScript, estilo, pruebas, compilación local y manifiesto. La segunda genera `dist/` con URLs de GitHub Pages, valida el manifiesto alojado y crea `dist/wordcodeaddin.zip`, con iconos y traducción al inglés. No instala nada en Word ni abre un servidor.

Los cambios en `main` se publican con el workflow **Publish GitHub Pages**. La publicación usa exclusivamente `dist/`; no sube node_modules, credenciales ni configuración local. Para alojar en otra dirección, defina `ADDIN_PRODUCTION_URL` con una URL HTTPS terminada en `/` antes de ejecutar `npm run build:hosted` y ajuste la dirección de la guía de instalación.

GitHub Pages es gratuito para este repositorio público, sujeto a sus [límites y condiciones](https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits). Esta publicación no es una entrada en la tienda de Office ni sustituye las políticas de instalación de Microsoft 365.

Consulte [privacidad](PRIVACY.md), [licencia MIT](LICENSE), [atribuciones](NOTICE.md) y [validación](VALIDATION.md).
