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
const { insertCode } = require("../src/taskpane/word.ts");

test("inserts at selection end, moves cursor and syncs on every insertion", async () => {
  const calls = [];
  global.Word = {
    InsertLocation: { end: "End" },
    SelectionMode: { end: "End" },
    run: async (callback) =>
      callback({
        document: {
          getSelection: () => ({
            insertOoxml: (xml, location) => {
              calls.push(["insert", location, xml.includes("Consolas")]);
              return { select: (mode) => calls.push(["select", mode]) };
            },
          }),
        },
        sync: async () => calls.push(["sync"]),
      }),
  };
  await insertCode("const x = 1;", "javascript", "light");
  await insertCode("print(1)", "python", "dark");
  assert.deepEqual(calls, [
    ["insert", "End", true],
    ["select", "End"],
    ["sync"],
    ["insert", "End", true],
    ["select", "End"],
    ["sync"],
  ]);
});

test("propagates Word errors so the panel can report them", async () => {
  global.Word = {
    run: async () => {
      throw new Error("Document is protected");
    },
  };
  await assert.rejects(insertCode("x", "plaintext", "light"), /protected/);
});
