import hljs from "highlight.js/lib/core";
import csharp from "highlight.js/lib/languages/csharp";
import python from "highlight.js/lib/languages/python";
import javascript from "highlight.js/lib/languages/javascript";
import typescript from "highlight.js/lib/languages/typescript";
import java from "highlight.js/lib/languages/java";
import xml from "highlight.js/lib/languages/xml";
import css from "highlight.js/lib/languages/css";
import sql from "highlight.js/lib/languages/sql";
import json from "highlight.js/lib/languages/json";
import bash from "highlight.js/lib/languages/bash";

const grammars = { csharp, python, javascript, typescript, java, xml, css, sql, json, bash };
Object.keys(grammars).forEach((name) => hljs.registerLanguage(name, grammars[name]));
export type CodeLanguage = keyof typeof grammars | "plaintext";
export type CodeTheme = "light" | "dark";
export const languages: [CodeLanguage, string][] = [
  ["csharp", "C#"],
  ["python", "Python"],
  ["javascript", "JavaScript"],
  ["typescript", "TypeScript"],
  ["java", "Java"],
  ["xml", "HTML"],
  ["css", "CSS"],
  ["sql", "SQL"],
  ["json", "JSON"],
  ["bash", "Bash"],
  ["plaintext", "Texto plano"],
];

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

const palettes = {
  light: {
    base: "24292F",
    keyword: "8250DF",
    string: "0A3069",
    comment: "57606A",
    number: "0550AE",
    title: "953800",
    background: "F6F8FA",
    border: "D0D7DE",
  },
  dark: {
    base: "E6EDF3",
    keyword: "FF7B72",
    string: "A5D6FF",
    comment: "ADB6C0",
    number: "79C0FF",
    title: "D2A8FF",
    background: "161B22",
    border: "57606A",
  },
};

// Only parse the library's span markup. Its HTML is never inserted into Word.
export function codeRuns(code: string, language: CodeLanguage): { text: string; kind: string }[] {
  if (!languages.some(([name]) => name === language)) throw new Error("Lenguaje no compatible.");
  const normalized = code.replace(/\r\n?/g, "\n");
  if (
    normalized.split("").some((character) => {
      const value = character.charCodeAt(0);
      return (value < 32 && value !== 9 && value !== 10) || value === 65534 || value === 65535;
    })
  ) {
    throw new Error("El código contiene caracteres de control no compatibles con Word.");
  }
  if (language === "plaintext") return [{ text: normalized, kind: "base" }];
  const html = hljs.highlight(normalized, { language, ignoreIllegals: true }).value;
  const stack: string[] = ["base"];
  const runs: { text: string; kind: string }[] = [];
  for (const part of html.split(/(<[^>]*>)/g)) {
    if (!part) continue;
    if (part === "</span>") {
      stack.pop();
      continue;
    }
    if (part.startsWith("<")) {
      const match = /^<span class="([a-zA-Z0-9_ -]+)">$/.exec(part);
      if (!match) throw new Error("Formato del resaltador no compatible.");
      const cls = match[1];
      stack.push(
        /comment/.test(cls)
          ? "comment"
          : /string|attr/.test(cls)
            ? "string"
            : /keyword|literal|type/.test(cls)
              ? "keyword"
              : /number/.test(cls)
                ? "number"
                : /title|built_in/.test(cls)
                  ? "title"
                  : stack[stack.length - 1]
      );
      continue;
    }
    const entities = {
      "&amp;": "&",
      "&lt;": "<",
      "&gt;": ">",
      "&quot;": '"',
      "&#x27;": "'",
      "&#39;": "'",
    };
    runs.push({
      text: part.replace(/&(?:amp|lt|gt|quot|#x27|#39);/g, (entity) => entities[entity]),
      kind: stack[stack.length - 1],
    });
  }
  if (runs.map((run) => run.text).join("") !== normalized)
    throw new Error("No se pudo conservar el texto del código.");
  return runs;
}

export function createCodeOoxml(code: string, language: CodeLanguage, theme: CodeTheme): string {
  if (!code.trim()) throw new Error("Pegue código antes de insertarlo.");
  if (theme !== "light" && theme !== "dark") throw new Error("Tema no compatible.");
  const palette = palettes[theme];
  const runs = codeRuns(code, language)
    .map(({ text, kind }) => {
      const content = text
        .split(/(\n|\t)/g)
        .map((part) =>
          part === "\n"
            ? "<w:br/>"
            : part === "\t"
              ? "<w:tab/>"
              : `<w:t xml:space="preserve">${escapeXml(part)}</w:t>`
        )
        .join("");
      return `<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:color w:val="${palette[kind]}"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>${content}</w:r>`;
    })
    .join("");
  const tabs = Array.from(
    { length: 40 },
    (_, index) => `<w:tab w:val="left" w:pos="${(index + 1) * 480}"/>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:p><w:pPr><w:pStyle w:val="Normal"/><w:numPr><w:numId w:val="0"/></w:numPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="${palette.border}"/><w:bottom w:val="single" w:sz="4" w:space="6" w:color="${palette.border}"/><w:left w:val="single" w:sz="4" w:space="6" w:color="${palette.border}"/><w:right w:val="single" w:sz="4" w:space="6" w:color="${palette.border}"/></w:pBdr><w:shd w:val="clear" w:fill="${palette.background}"/><w:tabs>${tabs}</w:tabs><w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/><w:ind w:left="120" w:right="120" w:firstLine="0"/></w:pPr>${runs}</w:p><w:p><w:pPr><w:pStyle w:val="Normal"/><w:pBdr><w:top w:val="nil"/><w:bottom w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/></w:pBdr><w:shd w:val="clear" w:fill="auto"/><w:spacing w:before="0" w:after="0"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="000000"/><w:sz w:val="22"/></w:rPr></w:pPr></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}
