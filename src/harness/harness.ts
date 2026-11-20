// Test harness (not shipped): renders every template through the same pipeline the add-in uses.
// Driven by scripts/render-all.mjs.
import { DEFAULT_OPTIONS, DiagramOptions } from "../options";
import { renderDiagram, renderPreview } from "../render/mermaidRenderer";
import { TEMPLATES } from "../templates";
import { decodeMeta, encodeMeta } from "../powerpoint/sourceMeta";
import { capitalizationRepairs, findShapeCodeBlock } from "../powerpoint/codeBlocks";
import { DEFAULT_SLIDE, fitInto, pxToPt, rescale, slideArea } from "../powerpoint/layout";
import { findIconRefs } from "../render/iconPacks";

interface HarnessResult {
  id: string;
  name: string;
  ok: boolean;
  error?: string;
  svg?: string | null;
  png?: string;
  meta?: string;
  width?: number;
  height?: number;
  /** Where Insert would put it on an empty 16:9 slide, in points. */
  box?: { left: number; top: number; width: number; height: number };
  foreignObject?: boolean;
  styleBlock?: boolean;
}

async function renderAll(overrides: Partial<DiagramOptions> = {}): Promise<HarnessResult[]> {
  const options = { ...DEFAULT_OPTIONS, ...overrides };
  const results: HarnessResult[] = [];
  for (const t of TEMPLATES) {
    try {
      const r = await renderDiagram(t.source, options);
      results.push({
        id: t.id,
        name: t.name,
        ok: true,
        svg: r.svg,
        png: r.pngBase64,
        width: r.width,
        height: r.height,
        box: fitInto(pxToPt(r), slideArea(DEFAULT_SLIDE), options.size),
        foreignObject: !!r.svg && r.svg.includes("<foreignObject"),
        styleBlock: !!r.svg && r.svg.includes("<style"),
        meta: encodeMeta(t.source, options, { width: r.width, height: r.height }),
      });
    } catch (err) {
      results.push({ id: t.id, name: t.name, ok: false, error: err instanceof Error ? err.message : String(err) });
    }
  }
  return results;
}

async function selfTest(): Promise<string[]> {
  const failures: string[] = [];
  const check = (cond: boolean, msg: string) => cond || failures.push(msg);
  const near = (a: number, b: number) => Math.abs(a - b) < 0.02;

  // Source metadata round trip, including unicode and line breaks mangled by the host.
  const src = 'flowchart TD\n  A["Ünïcødé ✓"] --> B\n  B --> C';
  const opts = { ...DEFAULT_OPTIONS, theme: "dark" as const, scale: 3, size: "actual" as const };
  const meta = encodeMeta(src, opts, { width: 120, height: 340 });
  const decoded = decodeMeta(meta);
  check(decoded?.source === src, "meta round trip: source");
  check(decoded?.options.theme === "dark" && decoded?.options.scale === 3, "meta round trip: options");
  check(decoded?.options.size === "actual", "meta round trip: size");
  check(decoded?.natural?.width === 120 && decoded?.natural?.height === 340, "meta round trip: natural size");
  check(decodeMeta(meta.replace(/\n/g, "\r"))?.source === src, "meta with CR line breaks");
  check(decodeMeta(meta.replace(/\n/g, "\u000B"))?.source === src, "meta with vertical-tab line breaks");
  check(decodeMeta("Some picture") === null, "non-mermaid text is ignored");

  // Alt text written by the Word add-in (no size fields) is readable.
  const wordAlt =
    '%% mermaid-word {"v":1,"theme":"forest","scale":2,"maxWidthIn":6.5,"background":"transparent","pngOnly":false,"src":"' +
    btoa("pie\n  \"A\": 1") +
    '"}\npie\n  "A": 1';
  const fromWord = decodeMeta(wordAlt);
  check(fromWord?.source === 'pie\n  "A": 1', "Word alt text: source");
  check(fromWord?.options.theme === "forest" && fromWord?.options.background === "transparent", "Word alt text: options");
  check(fromWord?.natural === undefined, "Word alt text: no natural size");

  // Code blocks in a text box: paragraphs (\r), soft breaks (\v) and AutoCorrect characters.
  const typed = findShapeCodeBlock("```mermaid\rflowchart LR\r  A → B[“Hi”]\u000B  B –– C\r```");
  check(typed?.source === 'flowchart LR\n  A --> B["Hi"]\n  B ---- C', `typed block: ${JSON.stringify(typed?.source)}`);
  check(typed?.hasOtherText === false, "typed block: no other text");
  const withText = findShapeCodeBlock("Our flow\r~~~mermaid\rpie\r~~~\r");
  check(withText?.source === "pie" && withText.hasOtherText, "block with other text");
  check(findShapeCodeBlock("```mermaid\rflowchart LR\r") === null, "unterminated block");
  check(findShapeCodeBlock("just text") === null, "no block");
  check(findShapeCodeBlock("```mermaid\rpie\r~~~") === null, "mismatched fences");

  // Undoing PowerPoint's auto-capitalization.
  const repairs = capitalizationRepairs("Flowchart TD\nStart --> End\nSubgraph X\nend");
  check(repairs[0] === "Flowchart TD".replace("F", "f") + "\nStart --> End\nSubgraph X\nend", `repair first line: ${repairs[0]}`);
  check(repairs[1] === "flowchart TD\nStart --> End\nsubgraph X\nend", `repair all lines: ${repairs[1]}`);
  check(capitalizationRepairs("SequenceDiagram\nParticipant Alice")[1] === "sequenceDiagram\nparticipant Alice", "repair sequence");
  check(capitalizationRepairs("C4Context\n  Person(a, \"A\")").length === 0, "C4 keywords untouched");
  check(capitalizationRepairs("flowchart TD\n  A --> B").length === 0, "nothing to repair");

  // The repaired source actually renders.
  try {
    await renderDiagram(capitalizationRepairs("Flowchart LR\n  A --> B")[0], DEFAULT_OPTIONS);
  } catch (err) {
    failures.push(`repaired source renders: ${err}`);
  }

  // Layout.
  const area = slideArea(DEFAULT_SLIDE);
  check(area.left === 27 && area.width === 906 && area.height === 486, `slide area: ${JSON.stringify(area)}`);
  const fit = fitInto({ width: 100, height: 50 }, area, "fit");
  check(near(fit.width, 906) && near(fit.height, 453) && near(fit.left, 27) && near(fit.top, 43.5), `fit: ${JSON.stringify(fit)}`);
  const actual = fitInto({ width: 100, height: 50 }, area, "actual");
  check(actual.width === 100 && near(actual.left, 430) && near(actual.top, 245), `actual: ${JSON.stringify(actual)}`);
  const tooBig = fitInto({ width: 2000, height: 500 }, area, "actual");
  check(near(tooBig.width, 906), `actual shrinks: ${JSON.stringify(tooBig)}`);
  // Re-render keeps the user's 2x scale and center.
  const kept = rescale({ left: 100, top: 100, width: 200, height: 100 }, { width: 100, height: 50 }, { width: 150, height: 50 }, DEFAULT_SLIDE);
  check(kept.width === 300 && kept.height === 100 && kept.left === 50 && kept.top === 100, `rescale: ${JSON.stringify(kept)}`);
  // ... but stays on the slide.
  const clamped = rescale({ left: 700, top: 400, width: 200, height: 100 }, { width: 100, height: 50 }, { width: 300, height: 100 }, DEFAULT_SLIDE);
  check(clamped.left + clamped.width <= 960 && clamped.top + clamped.height <= 540, `rescale clamp: ${JSON.stringify(clamped)}`);
  const shrunk = rescale({ left: 0, top: 0, width: 900, height: 300 }, { width: 300, height: 100 }, { width: 600, height: 100 }, DEFAULT_SLIDE);
  check(shrunk.width <= 960 && near(shrunk.width / shrunk.height, 6), `rescale shrink: ${JSON.stringify(shrunk)}`);

  // Iconify icon references (bundled "logos" and built-in names are skipped).
  const refs = findIconRefs(
    'architecture-beta\n  service a(mdi:database)[A]\n  service b(logos:aws)[B]\n  service c(cloud)[C]\n' +
      '  X@{ icon: "fa6-solid:user" }\n  style X stroke-width:4px'
  );
  const refList = JSON.stringify([...refs].map(([prefix, names]) => [prefix, [...names]]));
  check(refList === '[["mdi",["database"]],["fa6-solid",["user"]]]', `icon refs: ${refList}`);
  return failures;
}

async function renderRaw(id: string, overrides: Partial<DiagramOptions> = {}): Promise<string> {
  const t = TEMPLATES.find((x) => x.id === id)!;
  return renderPreview(t.source, { ...DEFAULT_OPTIONS, ...overrides });
}

Object.assign(window, { __renderAll: renderAll, __renderRaw: renderRaw, __selfTest: selfTest, __ready: true });
