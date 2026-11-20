/* global Office */
import { DEFAULT_OPTIONS, DiagramOptions, MermaidTheme } from "../options";
import { renderPreview, SupersededError } from "../render/mermaidRenderer";
import { TEMPLATES } from "../templates";
import {
  BatchResult,
  convertCodeBlocks,
  getSelectedDiagram,
  insertDiagram,
  refreshAllDiagrams,
  updateSelectedDiagram,
} from "../powerpoint/diagrams";

const STORAGE_KEY = "mermaid-powerpoint-state";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const ui = {
  template: $<HTMLSelectElement>("template"),
  source: $<HTMLTextAreaElement>("source"),
  error: $<HTMLDivElement>("error"),
  preview: $<HTMLDivElement>("preview"),
  theme: $<HTMLSelectElement>("theme"),
  background: $<HTMLSelectElement>("background"),
  size: $<HTMLSelectElement>("size"),
  scale: $<HTMLSelectElement>("scale"),
  pngOnly: $<HTMLInputElement>("pngOnly"),
  insert: $<HTMLButtonElement>("insert"),
  update: $<HTMLButtonElement>("update"),
  load: $<HTMLButtonElement>("load"),
  convert: $<HTMLButtonElement>("convert"),
  refresh: $<HTMLButtonElement>("refresh"),
  status: $<HTMLDivElement>("status"),
};

function readOptions(): DiagramOptions {
  return {
    theme: ui.theme.value as MermaidTheme,
    background: ui.background.value === "transparent" ? "transparent" : "white",
    size: ui.size.value === "actual" ? "actual" : "fit",
    scale: Number(ui.scale.value) || DEFAULT_OPTIONS.scale,
    pngOnly: ui.pngOnly.checked,
  };
}

function writeOptions(options: DiagramOptions) {
  ui.theme.value = options.theme;
  ui.background.value = options.background;
  ui.size.value = options.size;
  ui.scale.value = String(options.scale);
  ui.pngOnly.checked = options.pngOnly;
  syncScaleEnabled();
}

/** The PNG resolution only matters for PNG-only pictures: for SVG, PowerPoint makes its own fallback. */
function syncScaleEnabled() {
  ui.scale.disabled = !ui.pngOnly.checked;
  ui.scale.title = ui.pngOnly.checked ? "" : "Used with “Insert PNG only”";
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ source: ui.source.value, options: readOptions() }));
  } catch {
    // Storage can be unavailable in some hosts; persistence is optional.
  }
}

function loadState() {
  let source = TEMPLATES[0].source;
  let options = DEFAULT_OPTIONS;
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
    if (saved?.source) source = saved.source;
    if (saved?.options) options = { ...DEFAULT_OPTIONS, ...saved.options };
  } catch {
    // Ignore corrupt or unavailable storage.
  }
  ui.source.value = source;
  writeOptions(options);
}

function setStatus(message: string, isError = false) {
  ui.status.textContent = message;
  ui.status.classList.toggle("error-text", isError);
}

function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === "object" && err && "message" in err) return String((err as { message: unknown }).message);
  return String(err);
}

function describeBatch(result: BatchResult, verb: string): string {
  const lines = [`${verb} ${result.converted} diagram${result.converted === 1 ? "" : "s"}.`];
  for (const f of result.failures) lines.push(`✗ ${f.label}: ${f.error.split("\n")[0]}`);
  return lines.join("\n");
}

// ---- Preview -------------------------------------------------------------

let previewTimer: number | undefined;
let previewToken = 0;

function schedulePreview() {
  window.clearTimeout(previewTimer);
  previewTimer = window.setTimeout(updatePreview, 300);
}

async function updatePreview() {
  const token = ++previewToken;
  const source = ui.source.value;
  if (!source.trim()) {
    ui.error.hidden = true;
    ui.preview.innerHTML = '<div class="placeholder">Type Mermaid code above to see a preview.</div>';
    return;
  }
  try {
    const options = readOptions();
    const svg = await renderPreview(source, options);
    if (token !== previewToken) return;
    ui.preview.innerHTML = svg;
    ui.preview.classList.toggle("dark-theme", options.theme === "dark" && options.background === "transparent");
    ui.error.hidden = true;
  } catch (err) {
    if (token !== previewToken || err instanceof SupersededError) return;
    ui.error.textContent = errorMessage(err);
    ui.error.hidden = false;
  }
}

// ---- Actions -------------------------------------------------------------

/**
 * True while an action runs. Actions change the selection (each placed diagram is selected),
 * so selection checks must not re-enable buttons until the action has finished.
 */
let busy = false;

async function runAction(busyText: string, action: () => Promise<string | void>) {
  if (busy) return;
  busy = true;
  const buttons = [ui.insert, ui.update, ui.load, ui.convert, ui.refresh];
  buttons.forEach((b) => (b.disabled = true));
  setStatus(busyText);
  try {
    const message = await action();
    setStatus(message ?? "");
  } catch (err) {
    setStatus(errorMessage(err), true);
  } finally {
    busy = false;
    [ui.insert, ui.load, ui.convert, ui.refresh].forEach((b) => (b.disabled = false));
    // Update's state depends on the selection, which the action may have changed.
    void checkSelection();
  }
}

let selectionTimer: number | undefined;

async function checkSelection() {
  if (busy) return;
  try {
    const data = await getSelectedDiagram();
    if (busy) return;
    ui.update.disabled = !data;
    ui.update.title = data ? "Re-render the selected diagram with the code above" : "Select a Mermaid diagram first";
  } catch {
    if (!busy) ui.update.disabled = true;
  }
}

function wireEvents() {
  for (const t of TEMPLATES) {
    const option = document.createElement("option");
    option.value = t.id;
    option.textContent = t.name;
    ui.template.appendChild(option);
  }

  ui.template.addEventListener("change", () => {
    const template = TEMPLATES.find((t) => t.id === ui.template.value);
    if (!template) return;
    ui.source.value = template.source;
    ui.template.value = "";
    saveState();
    schedulePreview();
  });

  ui.source.addEventListener("input", () => {
    saveState();
    schedulePreview();
  });

  ui.source.addEventListener("keydown", (e) => {
    if (e.key !== "Tab" || e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;
    e.preventDefault();
    const { selectionStart: start, selectionEnd: end, value } = ui.source;
    ui.source.value = value.slice(0, start) + "    " + value.slice(end);
    ui.source.selectionStart = ui.source.selectionEnd = start + 4;
    ui.source.dispatchEvent(new Event("input"));
  });

  for (const el of [ui.theme, ui.background, ui.size, ui.scale, ui.pngOnly]) {
    el.addEventListener("change", () => {
      syncScaleEnabled();
      saveState();
      schedulePreview();
    });
  }

  ui.insert.addEventListener("click", () =>
    runAction("Rendering…", async () => {
      await insertDiagram(ui.source.value, readOptions());
      return "Diagram inserted.";
    })
  );

  ui.update.addEventListener("click", () =>
    runAction("Updating…", async () => {
      const updated = await updateSelectedDiagram(ui.source.value, readOptions());
      return updated ? "Diagram updated." : "Select a Mermaid diagram on a slide first.";
    })
  );

  ui.load.addEventListener("click", () =>
    runAction("Reading selection…", async () => {
      const data = await getSelectedDiagram();
      if (!data) return "No Mermaid diagram is selected. Click a diagram inserted by this add-in.";
      ui.source.value = data.source;
      writeOptions(data.options);
      saveState();
      schedulePreview();
      return "Loaded the selected diagram. Edit it, then click Update selected.";
    })
  );

  ui.convert.addEventListener("click", () =>
    runAction("Converting code blocks…", async () => {
      const result = await convertCodeBlocks(readOptions());
      if (!result.converted && !result.failures.length) {
        return "No ```mermaid blocks found. In a text box, put ```mermaid on the first line, the code below it, then ``` on the last line.";
      }
      return describeBatch(result, "Converted");
    })
  );

  ui.refresh.addEventListener("click", () =>
    runAction("Re-rendering all diagrams…", async () => {
      const result = await refreshAllDiagrams(readOptions());
      if (!result.converted && !result.failures.length) return "No Mermaid diagrams found in the presentation.";
      return describeBatch(result, "Re-rendered") + "\n(Current options were applied to every diagram.)";
    })
  );

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(checkSelection, 250);
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("This add-in only works in Microsoft PowerPoint.", true);
    return;
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    setStatus("This version of PowerPoint is too old for this add-in (PowerPointApi 1.5 is required).", true);
    return;
  }
  loadState();
  wireEvents();
  void updatePreview();
  void checkSelection();
});
