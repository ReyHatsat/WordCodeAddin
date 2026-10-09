# WordCodeAddin

[Español](README.md) | **English (United States)**

Insert editable code blocks with syntax highlighting into Microsoft Word. The interface follows Word's display language: formal Spanish or American English. Other languages use Spanish. Code is processed locally and is never executed.

## Install on Windows

You will need **Word for Windows with an active Microsoft 365 account**, internet access, and [Node.js 22.15 or later](https://nodejs.org/en/download). Your organization must allow custom Office add-ins. This project uses the unified manifest; [check Microsoft's compatibility requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) if Word does not load it.

1. Select **Code → Download ZIP** on [GitHub](https://github.com/ReyHatsat/WordCodeAddin), then extract the ZIP.
2. Open the extracted folder and double-click **Install.cmd**. This is only needed the first time.
3. Sign in using **the same Microsoft 365 account as Word**. Approve the local HTTPS development certificate if prompted.
4. Word opens. Open an editable document and select **Insert → Code → Insert Code**. In Spanish Word, select **Insertar → Código → Insertar código**.

If you prefer a terminal, open one in the folder containing `package.json` and run these commands **one at a time**:

```powershell
npm ci
npm run signin:browser
npm run start:desktop:word:browser
```

**Keep the local server running while using the add-in.** This is a local installation, rather than an Office Store installation. Each computer needs its own copy and server. GitHub does not host the running add-in.

## Open it again

Double-click **Start.cmd**, leave its window open, and open Word. Alternatively, open a terminal in the project folder and run:

```powershell
npm run dev-server
```

Leave that terminal open, launch Word, and use the ribbon button. The add-in is already registered; sign-in and installation are only needed on the first run. To stop the server, press **Ctrl+C** in its terminal. Close Word before stopping it.

## Use

Paste your code, select a language and a Light or Dark theme, place the cursor in Word, and select **Insert into Word**. If text is selected, the block is added after the selection. Existing text is preserved. The inserted code remains editable, with Consolas 10-point text, colors, background, borders, line breaks, spaces, and tabs.

Supported languages: C#, Python, JavaScript, TypeScript, Java, HTML, CSS, SQL, JSON, Bash, and Plain Text. The add-in does not translate the code. English Word uses American English; Spanish Word uses formal Spanish. Reopen the pane after changing Word's display language.

## If installation fails

- **Sign-in error:** confirm that Word and the browser use the same Microsoft 365 account. The commands above include a browser sign-in workaround for Windows `platform_broker_error`. It changes only its child processes and does not modify Windows or dependencies. This workaround was verified with Agents Toolkit CLI 1.1.17; future versions may require adjustments.
- **Certificate error:** run `npx office-addin-dev-certs install`, accept the certificate, and try again.
- **No ribbon button:** confirm that your Word version supports unified Office manifests and that your organization permits sideloading. Restart Word. Microsoft 365 permissions or unsupported clients cannot be fixed by downloading the repository again.
- **Blank pane or unavailable localhost:** start `npm run dev-server` in the project folder and leave it running. Port 3000 must be available.
- **WebView debugging dialog:** choose **Cancel**. Use the normal startup commands above, which disable debugging.

## Development and distribution

Run `npm run check` for TypeScript, lint, tests, production build, and manifest validation. Run `npm run build:dev` for a development build. `dist/` contains the built panel, icons, manifest, and English localization file. Build output is excluded from Git.

For hosted distribution, provide an HTTPS server, set `ADDIN_PRODUCTION_URL` to its URL with a trailing slash before building, update manifest domains, and distribute the app package through the Microsoft 365 channels supported by your organization. This repository is ready for local installation; it does not publish a hosted service or an Office Store listing. The unified manifest currently retains `devPreview` from the original project.

See [privacy information](PRIVACY.md) and the [MIT license](LICENSE). Office.js is loaded from Microsoft's CDN. Authentication and installation use Microsoft's services; the add-in itself adds no telemetry and does not send code snippets to a server.

Run `npm run package` after building to generate `dist/wordcodeaddin.zip`. This ZIP contains the manifest, installation icons, and English translation; the pane requires the HTTPS server described above.
