// Optional workaround for a Windows MSAL native-broker failure.
// Applied only to child processes launched by this file; never edits dependencies.
const Module = require("node:module");
const { spawn } = require("node:child_process");

if (process.env.WORDCODEADDIN_BROWSER_AUTH === "1") {
  const originalLoad = Module._load;
  Module._load = function (request, parent, isMain) {
    const exports = originalLoad.call(this, request, parent, isMain);
    if (request !== "@azure/msal-node-extensions") return exports;
    // MSAL falls back to its browser flow when the broker reports unavailable.
    // Keep the persistence/cache exports unchanged.
    return {
      ...exports,
      NativeBrokerPlugin: class BrowserOnlyBroker {
        get isBrokerAvailable() {
          return false;
        }
      },
    };
  };
}

if (require.main === module) {
  const commands = {
    signin: "npm run signin",
    start: "npm run start:desktop:word",
    check: "npx --yes --package @microsoft/m365agentstoolkit-cli atk auth list",
  };
  const command = commands[process.argv[2]];
  if (!command) {
    console.error("Uso: node scripts/office-browser-auth.cjs signin|start|check");
    process.exitCode = 1;
  } else {
    const preload = JSON.stringify(__filename);
    const child = spawn(process.env.ComSpec || "cmd.exe", ["/d", "/s", "/c", command], {
      stdio: "inherit",
      windowsHide: true,
      env: {
        ...process.env,
        WORDCODEADDIN_BROWSER_AUTH: "1",
        NODE_OPTIONS: [process.env.NODE_OPTIONS, "--require " + preload].filter(Boolean).join(" "),
      },
    });
    child.on("error", (error) => {
      console.error(error.message);
      process.exitCode = 1;
    });
    child.on("exit", (code) => {
      process.exitCode = code === null ? 1 : code;
    });
  }
}
