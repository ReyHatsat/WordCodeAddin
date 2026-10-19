# WordCodeAddin

[Español](README.md) | **English (United States)**

Editable code blocks with syntax highlighting for Microsoft Word. The interface follows Word's display language: formal Spanish or American English; other languages use Spanish.

**[Download and install the add-in](https://reyhatsat.github.io/WordCodeAddin/)**

## Install the hosted version

This version loads the pane from GitHub Pages. **No Node.js, Start.cmd, or local server is needed to use it.** Internet access, Word with unified manifest support, and a Microsoft 365 account whose organization permits custom apps are required. The manifest retains `devPreview` from the original project; compatibility depends on your Office version and channel.

1. Download [wordcodeaddin.zip](https://reyhatsat.github.io/WordCodeAddin/wordcodeaddin.zip). **Do not extract it.**
2. Save your documents and close Word.
3. Open Microsoft Teams using the same account as Word. Select **Apps → Manage your apps → Upload an app → Upload a custom app**.
4. Select the downloaded ZIP, select **Add**, and close Teams. You do not need to open the add-in in Teams.
5. Open Word and an editable document. Select **Insert → Code → Insert Code**. If it is missing, select WordCodeAddin from **Home → Add-ins**. It may take a few minutes to appear.

If custom app upload is unavailable, your administrator must enable it or distribute the package through your organization's supported Microsoft 365 channels. See [Microsoft's installation procedure](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest#sideload-through-the-teams-app-store) and [compatibility requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview).

**After installation, simply open Word.** The pane and icons are downloaded from GitHub Pages when Word needs them. The hosted address must remain available.

### If the local version is already installed

Publishing this website does not change your installation. To switch to the hosted version, install the ZIP above once; both versions use the same add-in identifier. The hosted version does not need a local server. If Word retains the old registration, close its windows and review Microsoft's cache update guidance before reinstalling. Do not delete documents or entire Office folders to update the add-in.

## Use

Paste your code, select a language and a Light or Dark theme, place the cursor in Word, and select **Insert into Word**. The block is added at the end of the selection without replacing existing text and remains editable.

Supports C#, Python, JavaScript, TypeScript, Java, HTML, CSS, SQL, JSON, Bash, and Plain Text, with Consolas 10-point text, colors, background, borders, spaces, line breaks, and tabs. The add-in does not translate or execute code; processing happens on your computer. Reopen the pane after changing Word's display language.

## Troubleshooting

- **Missing ribbon button:** confirm unified manifest support, the same signed-in account, and your organization's custom app permissions. Look for WordCodeAddin under Home → Add-ins and restart Word.
- **Upload a custom app is unavailable:** please contact your Microsoft 365 administrator. Downloading the project again does not change organizational policies.
- **Pane unavailable:** check your Internet connection and access to `reyhatsat.github.io`. The hosted version does not use localhost or development certificates.
- **Word still opens localhost:** the local version is still registered. Install the hosted package and restart Word. A GitHub update does not automatically update Word's registration.
- **Protected document:** allow editing or use an editable document.

## Local development (optional)

This section is for modifying the code. It requires [Node.js 22.15 or later](https://nodejs.org/en/download), Word for Windows, and permission to install custom add-ins.

Download the source using **Code → Download ZIP**, extract it, and double-click **Install.cmd**. Alternatively, run these commands in the folder containing `package.json`:

```powershell
npm ci
npm run signin:browser
npm run start:desktop:word:browser
```

Use the same account as Word. Approve the local HTTPS certificate if prompted. For subsequent development sessions, open **Start.cmd** or run `npm run dev-server` and keep the server running. These scripts install the local variant, rather than the hosted version.

Browser sign-in includes a workaround for Windows `platform_broker_error`, verified with Agents Toolkit CLI 1.1.17. It changes only its child processes, rather than Windows or dependencies. Toolkit updates may require adjustments. For local certificate errors, run `npx office-addin-dev-certs install`. Use normal startup without debugging to avoid the WebView debugging pause.

## Build and publish

```powershell
npm run check
npm run build:hosted
```

The first command runs TypeScript, lint, tests, a local build, and manifest validation. The second generates `dist/` with GitHub Pages URLs, validates the hosted manifest, and creates `dist/wordcodeaddin.zip` with icons and the English translation. It does not install into Word or open a server.

Changes to `main` publish through the **Publish GitHub Pages** workflow. Only `dist/` is deployed; node_modules, credentials, and local configuration are excluded. For another host, set `ADDIN_PRODUCTION_URL` to an HTTPS URL with a trailing `/` before running `npm run build:hosted`, and update the installation guide's address.

GitHub Pages is free for this public repository, subject to its [limits and terms](https://docs.github.com/en/pages/getting-started-with-github-pages/github-pages-limits). This deployment is not an Office Store listing and does not override Microsoft 365 installation policies.

See [privacy information](PRIVACY.md), the [MIT license](LICENSE), [attribution](NOTICE.md), and [validation](VALIDATION.md).
