# WordCodeAddin

**Español** | [English (United States)](README.en.md)

Inserte bloques de código editables con resaltado de sintaxis en Microsoft Word. La interfaz sigue el idioma de Word: español formal o inglés de Estados Unidos. Otros idiomas usan español. El código se procesa localmente y nunca se ejecuta.

## Instalar en Windows

Necesita **Word para Windows con una cuenta activa de Microsoft 365**, conexión a Internet y [Node.js 22.15 o posterior](https://nodejs.org/es/download). Su organización debe permitir complementos personalizados. Este proyecto usa el manifiesto unificado; [consulte los requisitos de Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) si Word no lo carga.

1. En [GitHub](https://github.com/ReyHatsat/WordCodeAddin), pulse **Code → Download ZIP** y extraiga el archivo ZIP.
2. Abra la carpeta extraída y haga doble clic en **Install.cmd**. Necesita hacerlo solo la primera vez.
3. Inicie sesión con **la misma cuenta de Microsoft 365 que utiliza en Word**. Acepte el certificado HTTPS local si se solicita.
4. Word se abrirá. Abra un documento editable y seleccione **Insertar → Código → Insertar código**. Si Word está en inglés, seleccione **Insert → Code → Insert Code**.

Si prefiere usar una terminal, ábrala en la carpeta que contiene `package.json` y ejecute estas órdenes **una por una**:

```powershell
npm ci
npm run signin:browser
npm run start:desktop:word:browser
```

**Mantenga el servidor local funcionando mientras utiliza el complemento.** Esta es una instalación local. Cada equipo necesita su propia copia y servidor. GitHub permite descargar el proyecto, pero no aloja el complemento en ejecución.

## Abrirlo de nuevo

Haga doble clic en **Start.cmd**, deje la ventana abierta y abra Word. También puede abrir una terminal en la carpeta del proyecto y ejecutar:

```powershell
npm run dev-server
```

Deje esa terminal abierta, abra Word y utilice el botón de la cinta. El complemento ya está registrado; el inicio de sesión y la instalación solo se necesitan la primera vez. Para detener el servidor, pulse **Ctrl+C** en su terminal. Cierre Word antes de detenerlo.

## Uso

Pegue su código, seleccione un lenguaje y el tema Claro u Oscuro, sitúe el cursor en Word y pulse **Insertar en Word**. Si hay texto seleccionado, el bloque se añade después de la selección. El texto existente se conserva. El código insertado queda editable, con fuente Consolas de 10 puntos, colores, fondo, bordes, saltos de línea, espacios y tabulaciones.

Lenguajes disponibles: C#, Python, JavaScript, TypeScript, Java, HTML, CSS, SQL, JSON, Bash y Texto plano. El complemento no traduce el código. Word en inglés utiliza inglés estadounidense; Word en español utiliza español formal. Vuelva a abrir el panel después de cambiar el idioma de la interfaz de Word.

## Si la instalación falla

- **Error al iniciar sesión:** compruebe que Word y el navegador utilizan la misma cuenta de Microsoft 365. Las órdenes anteriores incluyen una alternativa de autenticación por navegador para el error `platform_broker_error` de Windows. Solo modifica sus procesos hijos; no altera Windows ni las dependencias. Se verificó con Agents Toolkit CLI 1.1.17; futuras versiones pueden requerir ajustes.
- **Error de certificado:** ejecute `npx office-addin-dev-certs install`, acepte el certificado y vuelva a intentarlo.
- **No aparece el botón:** compruebe que su versión de Word admite manifiestos unificados y que su organización permite instalar complementos personalizados. Reinicie Word. Volver a descargar el repositorio no resuelve restricciones de su cuenta o un cliente incompatible.
- **Panel vacío o localhost no disponible:** ejecute `npm run dev-server` en la carpeta del proyecto y déjelo funcionando. El puerto 3000 debe estar disponible.
- **Advertencia de depuración de WebView:** pulse **Cancelar**. Utilice las órdenes normales de inicio indicadas arriba, que desactivan la depuración.

## Desarrollo y distribución

Ejecute `npm run check` para comprobar TypeScript, estilo, pruebas, compilación optimizada y manifiesto. Use `npm run build:dev` para compilar en modo desarrollo. `dist/` contiene el panel compilado, los iconos, el manifiesto y el archivo de traducción al inglés. La compilación se excluye de Git.

Para distribuir una versión alojada, proporcione un servidor HTTPS, defina `ADDIN_PRODUCTION_URL` con su dirección terminada en `/` antes de compilar, ajuste los dominios del manifiesto y distribuya el paquete por los canales de Microsoft 365 permitidos por su organización. Este repositorio está preparado para instalación local; no publica un servicio alojado ni una entrada en la tienda de Office. El manifiesto unificado conserva `devPreview` del proyecto original.

Consulte la [información de privacidad](PRIVACY.md) y la [licencia MIT](LICENSE). Office.js se carga del CDN de Microsoft. La autenticación y la instalación utilizan servicios de Microsoft; el complemento no añade telemetría ni envía fragmentos de código a un servidor.

Use `npm run package` después de compilar para generar `dist/wordcodeaddin.zip`. Este ZIP contiene el manifiesto, los iconos de instalación y la traducción al inglés; el panel necesita el servidor HTTPS indicado arriba.
