// Placement math for diagrams on a slide. All boxes are in points (1/72 inch), the unit
// PowerPoint's API uses. Pure functions so the test harness can check them.

import { SizeMode } from "../options";

export interface Size {
  width: number;
  height: number;
}

export interface Box extends Size {
  left: number;
  top: number;
}

/** Mermaid sizes are CSS pixels (96 per inch); PowerPoint uses points (72 per inch). */
export const PT_PER_PX = 0.75;

/** Default 16:9 slide, used when PowerPoint can't report the slide size (before PowerPointApi 1.10). */
export const DEFAULT_SLIDE: Size = { width: 960, height: 540 };

export function pxToPt(size: Size): Size {
  return { width: size.width * PT_PER_PX, height: size.height * PT_PER_PX };
}

/** The slide minus a margin of 5% of its shorter side. */
export function slideArea(slide: Size): Box {
  const margin = Math.round(Math.min(slide.width, slide.height) * 0.05);
  return { left: margin, top: margin, width: slide.width - 2 * margin, height: slide.height - 2 * margin };
}

/**
 * Sizes a diagram for `area` and centers it there. "fit" scales up or down to fill the area;
 * "actual" keeps the natural size and only shrinks a diagram that doesn't fit.
 */
export function fitInto(natural: Size, area: Box, mode: SizeMode): Box {
  let scale = Math.min(area.width / natural.width, area.height / natural.height);
  if (mode === "actual") scale = Math.min(1, scale);
  const width = natural.width * scale;
  const height = natural.height * scale;
  return round({
    left: area.left + (area.width - width) / 2,
    top: area.top + (area.height - height) / 2,
    width,
    height,
  });
}

/**
 * Box for a re-rendered diagram that replaces `old`. The user's scaling is kept (the new
 * picture is scaled by the same factor relative to its natural size) and it stays centered
 * where the old one was. If the old picture was on the slide, the new one is kept on it too.
 * Without the old natural size, the old width is kept.
 */
export function rescale(old: Box, oldNatural: Size | undefined, newNatural: Size, slide: Size): Box {
  const sx = oldNatural ? old.width / oldNatural.width : old.width / newNatural.width;
  const sy = oldNatural ? old.height / oldNatural.height : sx;
  let width = newNatural.width * sx;
  let height = newNatural.height * sy;
  let left = old.left + (old.width - width) / 2;
  let top = old.top + (old.height - height) / 2;

  const wasOnSlide =
    old.left >= -0.5 && old.top >= -0.5 && old.left + old.width <= slide.width + 0.5 && old.top + old.height <= slide.height + 0.5;
  if (wasOnSlide) {
    const shrink = Math.min(1, slide.width / width, slide.height / height);
    if (shrink < 1) {
      left += (width - width * shrink) / 2;
      top += (height - height * shrink) / 2;
      width *= shrink;
      height *= shrink;
    }
    left = Math.min(Math.max(left, 0), slide.width - width);
    top = Math.min(Math.max(top, 0), slide.height - height);
  }
  return round({ left, top, width, height });
}

function round(box: Box): Box {
  const r = (n: number) => Math.round(n * 100) / 100;
  return { left: r(box.left), top: r(box.top), width: r(box.width), height: r(box.height) };
}
