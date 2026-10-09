/* global Word */
import { createCodeOoxml, CodeLanguage, CodeTheme } from "./code-format";

export async function insertCode(code: string, language: CodeLanguage, theme: CodeTheme) {
  const ooxml = createCodeOoxml(code, language, theme);
  await Word.run(async (context) => {
    const inserted = context.document.getSelection().insertOoxml(ooxml, Word.InsertLocation.end);
    inserted.select(Word.SelectionMode.end);
    await context.sync();
  });
}
