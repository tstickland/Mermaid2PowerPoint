/* global PowerPoint, Office */
// Insert, update, re-render and convert operations using PowerPoint's JavaScript API.
//
// PowerPoint's API can't add a picture directly, so pictures go in through the Common API
// (setSelectedDataAsync, which inserts on the current slide). The new shape is then found by
// comparing the slide's shape IDs before and after, and tagged with the diagram's source.

import { DiagramOptions } from "../options";
import { renderDiagram, RenderResult } from "../render/mermaidRenderer";
import { capitalizationRepairs, findShapeCodeBlock } from "./codeBlocks";
import { Box, DEFAULT_SLIDE, fitInto, pxToPt, rescale, Size, slideArea } from "./layout";
import { DiagramData, decodeMeta, encodeMeta } from "./sourceMeta";

/** Shape tag that holds the diagram's source and options (PowerPoint upper-cases tag keys). */
export const DIAGRAM_TAG = "MERMAID_DIAGRAM";
export const DIAGRAM_NAME = "Mermaid diagram";

export interface BatchResult {
  converted: number;
  failures: { label: string; error: string }[];
}

interface Rendered {
  result: RenderResult;
  /** Encoded source and options, stored in the tag and alt text. */
  meta: string;
  /** Natural size in points. */
  natural: Size;
}

/** A diagram (or code block) found on a slide. */
interface Located {
  slideId: string;
  slideIndex: number;
  shapeId: string;
  box: Box;
}

const SHAPE_PROPS = "items/id,items/type,items/left,items/top,items/width,items/height";

const supports = (version: string) => Office.context.requirements.isSetSupported("PowerPointApi", version);
const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function boxOf(shape: PowerPoint.Shape): Box {
  return { left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}

// ---- Rendering ---------------------------------------------------------------

/** SVG goes in as-is (PowerPoint makes its own fallback), so a PNG is only needed without SVG. */
function needsPng(options: DiagramOptions): boolean {
  return options.pngOnly || !Office.context.requirements.isSetSupported("ImageCoercion", "1.2");
}

async function render(source: string, options: DiagramOptions): Promise<Rendered> {
  const result = await renderDiagram(source, options, { png: needsPng(options) });
  return {
    result,
    meta: encodeMeta(source, options, { width: result.width, height: result.height }),
    natural: pxToPt(result),
  };
}

/** Renders a code block typed in PowerPoint, undoing auto-capitalization if Mermaid rejects it. */
async function renderTyped(source: string, options: DiagramOptions): Promise<Rendered> {
  try {
    return await render(source, options);
  } catch (err) {
    for (const candidate of capitalizationRepairs(source)) {
      try {
        return await render(candidate, options);
      } catch {
        // Try the next variant.
      }
    }
    throw err;
  }
}

// ---- Reading the presentation ------------------------------------------------

async function getSlideSize(context: PowerPoint.RequestContext): Promise<Size> {
  if (supports("1.10")) {
    try {
      const setup = context.presentation.pageSetup;
      setup.load("slideWidth,slideHeight");
      await context.sync();
      if (setup.slideWidth > 0 && setup.slideHeight > 0) return { width: setup.slideWidth, height: setup.slideHeight };
    } catch {
      // Fall back to the default size.
    }
  }
  return DEFAULT_SLIDE;
}

/** Reads each shape's diagram data from its tag, or from alt text for pictures pasted from Word. */
async function readDiagramData(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[]
): Promise<(DiagramData | null)[]> {
  const tags = shapes.map((shape) => {
    const tag = shape.tags.getItemOrNullObject(DIAGRAM_TAG);
    tag.load("value");
    return tag;
  });
  await context.sync();
  const data = tags.map((tag) => (tag.isNullObject ? null : decodeMeta(tag.value)));

  const pictures = shapes.filter((s, i) => !data[i] && (s.type === "Image" || s.type === "Graphic"));
  if (pictures.length && supports("1.10")) {
    try {
      pictures.forEach((s) => s.load("altTextDescription"));
      await context.sync();
      shapes.forEach((s, i) => {
        if (!data[i] && pictures.includes(s)) data[i] = decodeMeta(s.altTextDescription);
      });
    } catch {
      // Alt text is only a fallback.
    }
  }
  return data;
}

/** Every slide's shapes, in slide order. */
async function loadAllShapes(context: PowerPoint.RequestContext) {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const collections = slides.items.map((slide) => slide.shapes.load(SHAPE_PROPS));
  await context.sync();
  return slides.items.flatMap((slide, slideIndex) =>
    collections[slideIndex].items.map((shape) => ({ slideId: slide.id, slideIndex, shape }))
  );
}

async function findSelectedDiagram(
  context: PowerPoint.RequestContext
): Promise<(Located & { data: DiagramData }) | null> {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load(SHAPE_PROPS);
  await context.sync();
  if (!shapes.items.length) return null;

  const data = await readDiagramData(context, shapes.items);
  const index = data.findIndex((d) => d);
  if (index === -1) return null;
  const shape = shapes.items[index];
  const slide = shape.getParentSlideOrNullObject();
  slide.load("id");
  await context.sync();
  if (slide.isNullObject) return null; // On a layout or master.
  return { slideId: slide.id, slideIndex: 0, shapeId: shape.id, box: boxOf(shape), data: data[index]! };
}

async function findAllDiagrams(context: PowerPoint.RequestContext): Promise<(Located & { data: DiagramData })[]> {
  const entries = await loadAllShapes(context);
  const data = await readDiagramData(
    context,
    entries.map((e) => e.shape)
  );
  return entries.flatMap((e, i) =>
    data[i]
      ? [{ slideId: e.slideId, slideIndex: e.slideIndex, shapeId: e.shape.id, box: boxOf(e.shape), data: data[i]! }]
      : []
  );
}

const TEXT_SHAPE_TYPES = new Set(["TextBox", "Placeholder", "GeometricShape"]);

async function readTexts(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]): Promise<(string | null)[]> {
  try {
    const ranges = shapes.map((s) => s.textFrame.textRange.load("text"));
    await context.sync();
    return ranges.map((r) => r.text);
  } catch {
    // Some shapes (such as a filled picture placeholder) have no text: read them one by one.
    const texts: (string | null)[] = [];
    for (const shape of shapes) {
      try {
        const range = shape.textFrame.textRange.load("text");
        await context.sync();
        texts.push(range.text);
      } catch {
        texts.push(null);
      }
    }
    return texts;
  }
}

/**
 * Where a new diagram goes on the slide: an empty content placeholder if there is one (it is
 * replaced, as when you insert a picture into it), else the slide below its title, else the
 * whole slide. Placeholder types need PowerPointApi 1.8.
 */
async function freeArea(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  size: Size
): Promise<{ area: Box; placeholderId?: string }> {
  const area = slideArea(size);
  if (!supports("1.8")) return { area };
  try {
    const shapes = slide.shapes.load(SHAPE_PROPS);
    await context.sync();
    const placeholders = shapes.items.filter((s) => s.type === "Placeholder");
    const formats = placeholders.map((s) => s.placeholderFormat.load("type"));
    await context.sync();
    const typeOf = (s: PowerPoint.Shape) => formats[placeholders.indexOf(s)].type;

    const content = placeholders.filter((s) => typeOf(s) === "Body" || typeOf(s) === "Content");
    const texts = await readTexts(context, content);
    const empty = content
      .filter((_, i) => texts[i] !== null && !texts[i]!.trim())
      .sort((a, b) => a.top - b.top || a.left - b.left)[0];
    if (empty) return { area: boxOf(empty), placeholderId: empty.id };

    const title = placeholders.find((s) => typeOf(s) === "Title" || typeOf(s) === "CenterTitle");
    if (title && title.top + title.height < size.height / 2) {
      const top = title.top + title.height + area.top / 2;
      return { area: { ...area, top, height: area.top + area.height - top } };
    }
  } catch {
    // Use the whole slide.
  }
  return { area };
}

// ---- Writing -----------------------------------------------------------------

function setSelectedData(data: string, options: Office.SetSelectedDataOptions): Promise<void> {
  return new Promise((resolve, reject) =>
    Office.context.document.setSelectedDataAsync(data, options, (result) =>
      result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result.error.message))
    )
  );
}

/** Inserts the picture on the current slide: SVG (PowerPoint adds its own PNG fallback), or PNG. */
function insertPicture(result: RenderResult, box: Box): Promise<void> {
  const svg = !!result.svg && Office.context.requirements.isSetSupported("ImageCoercion", "1.2");
  if (!svg && !result.pngBase64) throw new Error("No PNG was rendered for this diagram.");
  return setSelectedData(svg ? result.svg! : result.pngBase64, {
    coercionType: svg ? Office.CoercionType.XmlSvg : Office.CoercionType.Image,
    imageLeft: box.left,
    imageTop: box.top,
    imageWidth: box.width,
    imageHeight: box.height,
  });
}

/**
 * Inserts a rendered diagram on a slide at `box`, tags it with its source and selects it.
 * With `replaceId`, that shape is deleted and the diagram takes its place in the z-order.
 */
async function placeDiagram(slideId: string, box: Box, rendered: Rendered, replaceId?: string): Promise<void> {
  // Make the slide current and clear the shape selection, so the picture lands on that slide
  // as a new shape rather than filling a selected placeholder.
  const { before, shapesAbove } = await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    const shapes = context.presentation.slides.getItem(slideId).shapes.load("items/id");
    await context.sync();
    try {
      context.presentation.slides.getItem(slideId).setSelectedShapes([]);
      await context.sync();
    } catch {
      // Not needed for the insertion to work.
    }
    let shapesAbove = 0;
    if (replaceId && supports("1.8")) {
      try {
        shapes.load("items/id,items/zOrderPosition");
        await context.sync();
        const old = shapes.items.find((s) => s.id === replaceId);
        if (old) shapesAbove = shapes.items.filter((s) => s.zOrderPosition > old.zOrderPosition).length;
      } catch {
        // The new diagram stays on top.
      }
    }
    return { before: new Set(shapes.items.map((s) => s.id)), shapesAbove };
  });

  await insertPicture(rendered.result, box);

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    let added: PowerPoint.Shape | undefined;
    for (let attempt = 0; attempt < 10 && !added; attempt++) {
      if (attempt) await delay(150);
      const shapes = slide.shapes.load("items/id");
      await context.sync();
      added = shapes.items.find((s) => !before.has(s.id));
    }
    if (!added) {
      throw new Error("The picture was inserted, but PowerPoint didn't report it, so it isn't editable.");
    }

    added.name = DIAGRAM_NAME;
    added.tags.add(DIAGRAM_TAG, rendered.meta);
    added.left = box.left;
    added.top = box.top;
    added.width = box.width;
    added.height = box.height;
    if (replaceId) slide.shapes.getItem(replaceId).delete();
    await context.sync();

    if (shapesAbove > 0) {
      try {
        for (let i = 0; i < shapesAbove; i++) added.setZOrder("SendBackward");
        await context.sync();
      } catch {
        // The diagram is already in place, just on top of the stack.
      }
    }
    if (supports("1.10")) {
      try {
        added.altTextTitle = DIAGRAM_NAME;
        added.altTextDescription = rendered.meta;
        await context.sync();
      } catch {
        // The tag already holds the source.
      }
    }
    try {
      slide.setSelectedShapes([added.id]);
      await context.sync();
    } catch {
      // Selection is a convenience.
    }
  });
}

// ---- Operations --------------------------------------------------------------

/** Inserts a new diagram on the current slide. */
export async function insertDiagram(source: string, options: DiagramOptions): Promise<void> {
  const rendered = await render(source, options);
  const target = await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides().load("items/id");
    await context.sync();
    if (!slides.items.length) throw new Error("Select a slide to insert the diagram on.");
    const slide = slides.items[0];
    const size = await getSlideSize(context);
    return { slideId: slide.id, ...(await freeArea(context, slide, size)) };
  });
  const box = fitInto(rendered.natural, target.area, options.size);
  await placeDiagram(target.slideId, box, rendered, target.placeholderId);
}

/** Returns the source and options of the selected diagram, or null. */
export async function getSelectedDiagram(): Promise<DiagramData | null> {
  return PowerPoint.run(async (context) => (await findSelectedDiagram(context))?.data ?? null);
}

/** Re-renders the selected diagram with new source and options. Returns false if none is selected. */
export async function updateSelectedDiagram(source: string, options: DiagramOptions): Promise<boolean> {
  const found = await PowerPoint.run(async (context) => {
    const diagram = await findSelectedDiagram(context);
    return diagram && { diagram, size: await getSlideSize(context) };
  });
  if (!found) return false;
  const { diagram, size } = found;
  const rendered = await render(source, options);
  const oldNatural = diagram.data.natural && pxToPt(diagram.data.natural);
  await placeDiagram(diagram.slideId, rescale(diagram.box, oldNatural, rendered.natural, size), rendered, diagram.shapeId);
  return true;
}

/** Puts the user back on the slides they had selected before a batch operation. */
async function restoreSlides(slideIds: string[]) {
  if (!slideIds.length) return;
  try {
    await PowerPoint.run(async (context) => {
      context.presentation.setSelectedSlides(slideIds);
      await context.sync();
    });
  } catch {
    // The slides may have been deleted meanwhile.
  }
}

async function selectedSlideIds(context: PowerPoint.RequestContext): Promise<string[]> {
  const slides = context.presentation.getSelectedSlides().load("items/id");
  await context.sync();
  return slides.items.map((s) => s.id);
}

/** Re-renders every Mermaid diagram in the presentation with the given options. */
export async function refreshAllDiagrams(options: DiagramOptions): Promise<BatchResult> {
  const { diagrams, size, selected } = await PowerPoint.run(async (context) => ({
    selected: await selectedSlideIds(context),
    diagrams: await findAllDiagrams(context),
    size: await getSlideSize(context),
  }));

  const result: BatchResult = { converted: 0, failures: [] };
  for (const d of diagrams) {
    try {
      const rendered = await render(d.data.source, options);
      const oldNatural = d.data.natural && pxToPt(d.data.natural);
      await placeDiagram(d.slideId, rescale(d.box, oldNatural, rendered.natural, size), rendered, d.shapeId);
      result.converted++;
    } catch (err) {
      result.failures.push({ label: `Slide ${d.slideIndex + 1}`, error: errorMessage(err) });
    }
  }
  if (diagrams.length) await restoreSlides(selected);
  return result;
}

/**
 * Replaces each text box that holds a ```mermaid fenced block with its rendered diagram,
 * sized to fit where the text box was.
 */
export async function convertCodeBlocks(options: DiagramOptions): Promise<BatchResult> {
  const { blocks, selected } = await PowerPoint.run(async (context) => {
    const selected = await selectedSlideIds(context);
    const entries = (await loadAllShapes(context)).filter((e) => TEXT_SHAPE_TYPES.has(e.shape.type));
    const texts = await readTexts(
      context,
      entries.map((e) => e.shape)
    );
    const blocks = entries.flatMap((e, i) => {
      const block = texts[i] === null ? null : findShapeCodeBlock(texts[i]!);
      return block
        ? [{ slideId: e.slideId, slideIndex: e.slideIndex, shapeId: e.shape.id, box: boxOf(e.shape), block }]
        : [];
    });
    return { blocks, selected };
  });

  const result: BatchResult = { converted: 0, failures: [] };
  for (const b of blocks) {
    const firstLine = b.block.source.split("\n").find((l) => l.trim())?.trim() ?? "(empty)";
    const label = `Slide ${b.slideIndex + 1}: ${firstLine}`;
    if (b.block.hasOtherText) {
      result.failures.push({
        label,
        error: "The text box has other text as well. Put the ```mermaid block in a text box of its own.",
      });
      continue;
    }
    try {
      const rendered = await renderTyped(b.block.source, options);
      await placeDiagram(b.slideId, fitInto(rendered.natural, b.box, "fit"), rendered, b.shapeId);
      result.converted++;
    } catch (err) {
      result.failures.push({ label, error: errorMessage(err) });
    }
  }
  if (blocks.length) await restoreSlides(selected);
  return result;
}
