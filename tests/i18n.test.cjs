const assert = require("node:assert/strict");
const { test } = require("node:test");
const ts = require("typescript");
const fs = require("node:fs");
require.extensions[".ts"] = (module, filename) => {
  module._compile(
    ts.transpileModule(fs.readFileSync(filename, "utf8"), {
      compilerOptions: {
        module: ts.ModuleKind.CommonJS,
        esModuleInterop: true,
        target: ts.ScriptTarget.ES2015,
      },
    }).outputText,
    filename
  );
};
const {
  getLocale,
  getStrings,
  spanish,
  english,
  errorMessage,
} = require("../src/taskpane/i18n.ts");

test("Word display language selects American English with Spanish fallback", () => {
  for (const locale of ["en-US", "EN-us", "en-GB", "en"]) {
    assert.equal(getLocale(locale), "en-US");
    assert.equal(getStrings(locale), english);
  }
  for (const locale of ["es-ES", "es-GT", "fr-FR", "", undefined]) {
    assert.equal(getLocale(locale), "es-ES");
    assert.equal(getStrings(locale), spanish);
  }
  assert.deepEqual(Object.keys(english).sort(), Object.keys(spanish).sort());
});

test("English errors translate formatter failures and handle unknown Office errors", () => {
  assert.equal(
    errorMessage(new Error("Tema no compatible."), "en-US"),
    "The selected theme is not supported."
  );
  assert.equal(errorMessage(new Error("Documento protegido"), "en-US"), english.failed);
  assert.equal(errorMessage(null, "en-US"), english.failed);
  assert.equal(errorMessage(new Error("Documento protegido"), "es-GT"), "Documento protegido");
});

test("manifest localization paths resolve and all translatable panel fields exist", () => {
  const manifest = require("../appPackage/manifest.json");
  const localized = require("../appPackage/en-us.json");
  assert.deepEqual(manifest.localizationInfo.additionalLanguages, [
    { languageTag: "en-us", file: "en-us.json" },
  ]);
  for (const [key, value] of Object.entries(localized)) {
    if (key === "$schema") continue;
    const path = key.replace(/\[(\d+)\]/g, ".$1").split(".");
    assert.equal(typeof path.reduce((object, segment) => object[segment], manifest), "string", key);
    assert.ok(value.trim(), key);
  }
  const html = fs.readFileSync(require.resolve("../src/taskpane/taskpane.html"), "utf8");
  for (const [, key] of html.matchAll(/data-i18n="([^"]+)"/g)) assert.ok(english[key], key);
});

test("English Word localizes the pane and insertion feedback without changing code", async () => {
  const Module = require("node:module");
  const originalLoad = Module._load;
  let submitted;
  let insertedCode;
  let officeReady;
  const fields = Object.fromEntries(
    ["code-form", "code", "language", "theme", "insert", "status"].map((id) => [
      id,
      {
        dataset: {},
        value: id === "language" ? "javascript" : id === "theme" ? "light" : "",
        setAttribute() {},
        addEventListener(event, callback) {
          submitted = callback;
        },
        appendChild(option) {
          if (option.value === "plaintext") this.plaintext = option;
        },
        querySelector() {
          return this.plaintext;
        },
        focus() {},
      },
    ])
  );
  const translated = Object.keys(english).map((key) => ({ dataset: { i18n: key } }));
  global.document = {
    documentElement: {},
    getElementById: (id) => fields[id],
    createElement: () => ({}),
    querySelectorAll: () => translated,
  };
  global.Office = {
    HostType: { Word: "Word" },
    context: { displayLanguage: "en-US", requirements: { isSetSupported: () => true } },
    onReady: (callback) => {
      officeReady = callback;
    },
  };
  Module._load = function (request, parent, isMain) {
    if (request === "./word" && parent.filename.endsWith("taskpane.ts")) {
      return {
        insertCode: async (code) => {
          insertedCode = code;
        },
      };
    }
    return originalLoad.call(this, request, parent, isMain);
  };
  try {
    require("../src/taskpane/taskpane.ts");
    officeReady({ host: "Word" });
    assert.equal(document.documentElement.lang, "en-US");
    assert.equal(document.title, "Insert Code");
    assert.equal(fields.code.placeholder, english.placeholder);
    assert.equal(fields.language.plaintext.textContent, "Plain Text");
    for (const element of translated)
      assert.equal(element.textContent, english[element.dataset.i18n]);
    assert.equal(fields.status.textContent, english.ready);
    await submitted({ preventDefault() {} });
    assert.equal(fields.status.textContent, english.empty);
    fields.code.value = "const saludo = 'Hola';\n";
    await submitted({ preventDefault() {} });
    assert.equal(insertedCode, fields.code.value);
    assert.equal(fields.status.textContent, english.success);
    assert.equal(fields.insert.disabled, false);
  } finally {
    Module._load = originalLoad;
    delete global.document;
    delete global.Office;
  }
});
