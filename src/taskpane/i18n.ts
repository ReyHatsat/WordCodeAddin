export const spanish = {
  title: "Insertar código",
  instructions:
    "Pegue su código y sitúe el cursor en el documento. Si hay texto seleccionado, el bloque se añade al final de la selección sin borrarlo.",
  code: "Código",
  placeholder: "Pegue aquí su código…",
  language: "Lenguaje",
  theme: "Tema",
  light: "Claro",
  dark: "Oscuro",
  plaintext: "Texto plano",
  insert: "Insertar en Word",
  initial: "Abra este panel desde Word para insertar código.",
  note: "El código se procesa localmente y queda editable. Las tabulaciones se conservan con posiciones cada cuatro caracteres.",
  ready: "Listo para insertar.",
  unsupported: "Se necesita Word con soporte de WordApi 1.1.",
  empty: "Error: pegue código antes de insertarlo.",
  inserting: "Insertando…",
  success: "Código insertado. Puede editarlo en el documento o insertar otro bloque.",
  failed: "No se pudo completar la inserción.",
  retry: "Compruebe que el documento permite editar, sitúe el cursor y vuelva a intentarlo.",
};

export const english: typeof spanish = {
  title: "Insert Code",
  instructions:
    "Paste your code and place the cursor in the document. If text is selected, the code block is added at the end of the selection without deleting it.",
  code: "Code",
  placeholder: "Paste your code here…",
  language: "Language",
  theme: "Theme",
  light: "Light",
  dark: "Dark",
  plaintext: "Plain Text",
  insert: "Insert into Word",
  initial: "Open this pane in Word to insert code.",
  note: "Code is processed locally and remains editable. Tabs are preserved with tab stops every four characters.",
  ready: "Ready to insert.",
  unsupported: "Word with WordApi 1.1 support is required.",
  empty: "Error: Please paste code before inserting it.",
  inserting: "Inserting…",
  success: "Code inserted. You may edit it in the document or insert another block.",
  failed: "The code could not be inserted.",
  retry: "Please confirm that the document allows editing, place the cursor, and try again.",
};

export function getLocale(displayLanguage = ""): "es-ES" | "en-US" {
  return /^en(?:[-_]|$)/i.test(displayLanguage) ? "en-US" : "es-ES";
}

export function getStrings(displayLanguage = "") {
  return getLocale(displayLanguage) === "en-US" ? english : spanish;
}

const errorTranslations: Record<string, string> = {
  "Lenguaje no compatible.": "The selected language is not supported.",
  "El código contiene caracteres de control no compatibles con Word.":
    "The code contains control characters that Word does not support.",
  "Formato del resaltador no compatible.": "The syntax highlighting format is not supported.",
  "No se pudo conservar el texto del código.": "The original code text could not be preserved.",
  "Pegue código antes de insertarlo.": "Please paste code before inserting it.",
  "Tema no compatible.": "The selected theme is not supported.",
};

export function errorMessage(error: unknown, displayLanguage = ""): string {
  const strings = getStrings(displayLanguage);
  if (!(error instanceof Error)) return strings.failed;
  if (getLocale(displayLanguage) === "en-US")
    return errorTranslations[error.message] || strings.failed;
  return error.message || strings.failed;
}
