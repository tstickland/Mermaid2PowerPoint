// The Mermaid source and render options are stored with each picture, in the same format the
// Word add-in (Mermaid2Word) uses for alt text:
//
//   %% mermaid-word {"v":1,"theme":"default",...,"src":"<base64 utf-8 source>"}
//   flowchart TD
//     A --> B
//
// `%%` lines are Mermaid comments, so the text stays readable and valid Mermaid. The base64
// copy guards against hosts that mangle line breaks. Sharing the Word format (marker included)
// means a diagram copied between Word and PowerPoint stays editable in either add-in.
// PowerPoint-only fields (`size`, and the natural size `w`/`h` in CSS pixels) are ignored by Word.

import { DEFAULT_OPTIONS, DiagramOptions } from "../options";

const MARKER = "%% mermaid-word ";

export interface DiagramData {
  source: string;
  options: DiagramOptions;
  /** Natural size of the rendered diagram in CSS pixels, when known. */
  natural?: { width: number; height: number };
}

function toBase64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

function fromBase64(b64: string): string {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

export function encodeMeta(
  source: string,
  options: DiagramOptions,
  natural: { width: number; height: number }
): string {
  const meta = { v: 1, ...options, w: natural.width, h: natural.height, src: toBase64(source) };
  return `${MARKER}${JSON.stringify(meta)}\n${source}`;
}

export function decodeMeta(text: string | null | undefined): DiagramData | null {
  if (!text) return null;
  const normalized = text.replace(/\r\n?|\u000B/g, "\n").replace(/^\s+/, "");
  if (!normalized.startsWith(MARKER.trim())) return null;

  const firstBreak = normalized.indexOf("\n");
  const header = firstBreak === -1 ? normalized : normalized.slice(0, firstBreak);
  const jsonStart = header.indexOf("{");
  const jsonEnd = header.lastIndexOf("}");
  if (jsonStart === -1 || jsonEnd === -1) return null;

  try {
    const meta = JSON.parse(header.slice(jsonStart, jsonEnd + 1));
    const source = meta.src ? fromBase64(meta.src) : firstBreak === -1 ? "" : normalized.slice(firstBreak + 1);
    const options: DiagramOptions = {
      theme: meta.theme ?? DEFAULT_OPTIONS.theme,
      scale: Number(meta.scale) || DEFAULT_OPTIONS.scale,
      background: meta.background === "transparent" ? "transparent" : "white",
      pngOnly: !!meta.pngOnly,
      size: meta.size === "actual" ? "actual" : "fit",
    };
    const w = Number(meta.w);
    const h = Number(meta.h);
    return { source, options, natural: w > 0 && h > 0 ? { width: w, height: h } : undefined };
  } catch {
    return null;
  }
}
