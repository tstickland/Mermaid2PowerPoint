export type MermaidTheme = "default" | "neutral" | "dark" | "forest" | "base";

/** How a newly inserted diagram is sized. */
export type SizeMode = "fit" | "actual";

export interface DiagramOptions {
  theme: MermaidTheme;
  /** Resolution multiplier for the PNG (used in PNG-only mode). */
  scale: number;
  background: "white" | "transparent";
  /** Skip the SVG and insert only the PNG. */
  pngOnly: boolean;
  /**
   * "fit": scale the diagram up or down to fill the free area of the slide.
   * "actual": insert at its natural size (96 px per inch), shrinking it only if it doesn't fit.
   */
  size: SizeMode;
}

export const DEFAULT_OPTIONS: DiagramOptions = {
  theme: "default",
  scale: 2,
  background: "white",
  pngOnly: false,
  size: "fit",
};
