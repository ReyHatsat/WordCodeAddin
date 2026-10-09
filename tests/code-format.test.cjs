const assert = require("node:assert/strict");
const { test } = require("node:test");
const ts = require("typescript");
const fs = require("node:fs");
// Transpile only our pure formatter in memory; no Word host or build output required.
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
  escapeXml,
  codeRuns,
  createCodeOoxml,
  languages,
} = require("../src/taskpane/code-format.ts");

function documentText(xml) {
  const entities = { "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'" };
  return [...xml.matchAll(/<w:t xml:space="preserve">([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/g)]
    .map((match) =>
      match[0] === "<w:tab/>"
        ? "\t"
        : match[0] === "<w:br/>"
          ? "\n"
          : match[1].replace(/&(amp|lt|gt|quot|apos);/g, (entity) => entities[entity])
    )
    .join("");
}

test("escapes markup, quotes and entity-looking input", () => {
  assert.equal(escapeXml("<&>\"'"), "&lt;&amp;&gt;&quot;&apos;");
  const input = '<script>alert("x")</script> &lt; </w:t><w:r> ñ 😀';
  for (const [language] of languages) {
    assert.equal(
      codeRuns(input, language)
        .map((r) => r.text)
        .join(""),
      input
    );
    const output = createCodeOoxml(input, language, "light");
    assert.ok(!output.includes("<script>"));
    assert.equal(documentText(output), input);
  }
});

test("preserves indentation, tabs, blank lines, trailing spaces and final newline", () => {
  const input = "  if (true) {\r\n\t  value();  \r\n\r\n}\r\n";
  for (const [language] of languages) {
    assert.equal(
      codeRuns(input, language)
        .map((r) => r.text)
        .join(""),
      input.replace(/\r\n/g, "\n")
    );
    const output = createCodeOoxml(input, language, "light");
    assert.equal(documentText(output), input.replace(/\r\n/g, "\n"));
    assert.equal((output.match(/<w:br\/>/g) || []).length, 4);
    assert.equal((output.match(/<w:tab\/>/g) || []).length, 1);
    assert.ok(output.includes('xml:space="preserve"'));
  }
});

test("C#, Python and JavaScript produce colors in both themes; plain text stays plain", () => {
  for (const [language, source] of [
    ["csharp", "public int x = 42;"],
    ["python", "def greet():\n    return 42"],
    ["javascript", "const x = 42;"],
  ]) {
    assert.ok(codeRuns(source, language).some((r) => r.kind === "keyword"));
    for (const theme of ["light", "dark"]) {
      const output = createCodeOoxml(source, language, theme);
      assert.ok(output.includes('w:ascii="Consolas"'));
      assert.ok(output.includes('w:val="20"'));
      assert.ok(output.includes(theme === "light" ? "8250DF" : "FF7B72"));
    }
  }
  assert.deepEqual(codeRuns("const x = 42;", "plaintext"), [
    { text: "const x = 42;", kind: "base" },
  ]);
});

test("rejects empty input, invalid options and XML control characters", () => {
  assert.throws(() => createCodeOoxml(" \n\t", "python", "light"));
  assert.throws(() => createCodeOoxml("x", "unknown", "light"));
  assert.throws(() => createCodeOoxml("x", "python", "unknown"));
  assert.throws(() => createCodeOoxml("x\u0000", "plaintext", "light"));
});
