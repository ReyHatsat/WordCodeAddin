/* global document, Office, HTMLFormElement, HTMLTextAreaElement, HTMLSelectElement, HTMLButtonElement, HTMLElement, HTMLOptionElement */
import { languages, CodeLanguage, CodeTheme } from "./code-format";
import { insertCode } from "./word";
import { getLocale, getStrings, errorMessage } from "./i18n";

const form = document.getElementById("code-form") as HTMLFormElement;
const code = document.getElementById("code") as HTMLTextAreaElement;
const language = document.getElementById("language") as HTMLSelectElement;
const theme = document.getElementById("theme") as HTMLSelectElement;
const button = document.getElementById("insert") as HTMLButtonElement;
const status = document.getElementById("status");
let displayLanguage = "es-ES";
let strings = getStrings(displayLanguage);
let ready = false;
let busy = false;

for (const [value, label] of languages) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  language.appendChild(option);
}

function message(text: string, error = false) {
  status.textContent = text;
  status.dataset.error = String(error);
}

Office.onReady((info) => {
  displayLanguage = Office.context.displayLanguage || "es-ES";
  strings = getStrings(displayLanguage);
  document.documentElement.lang = getLocale(displayLanguage);
  document.title = strings.title;
  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach((element) => {
    element.textContent = strings[element.dataset.i18n as keyof typeof strings];
  });
  code.placeholder = strings.placeholder;
  language.querySelector<HTMLOptionElement>('option[value="plaintext"]').textContent =
    strings.plaintext;
  ready =
    info.host === Office.HostType.Word &&
    Office.context.requirements.isSetSupported("WordApi", "1.1");
  button.disabled = !ready;
  message(ready ? strings.ready : strings.unsupported, !ready);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ready || busy) return;
  if (!code.value.trim()) {
    message(strings.empty, true);
    code.focus();
    return;
  }
  busy = true;
  button.disabled = true;
  form.setAttribute("aria-busy", "true");
  message(strings.inserting);
  try {
    await insertCode(code.value, language.value as CodeLanguage, theme.value as CodeTheme);
    message(strings.success);
  } catch (error) {
    const detail = errorMessage(error, displayLanguage);
    message(`Error: ${detail} ${strings.retry}`, true);
  } finally {
    busy = false;
    button.disabled = !ready;
    form.setAttribute("aria-busy", "false");
  }
});
