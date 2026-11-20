// Finds ```mermaid fenced code blocks typed into PowerPoint text boxes, and undoes the
// AutoCorrect changes PowerPoint makes while typing.

// PowerPoint's AutoCorrect replaces characters that Mermaid needs.
export function normalizeTypedText(line: string): string {
  return line
    .replace(/ /g, " ")
    .replace(/[“”„‟]/g, '"')
    .replace(/[‘’‚‛]/g, "'")
    .replace(/—/g, "---")
    .replace(/–/g, "--")
    .replace(/…/g, "...")
    .replace(/→/g, "-->")
    .replace(/←/g, "<--")
    .replace(/⇒/g, "==>")
    .replace(/⇐/g, "<==")
    .replace(/⇔/g, "<=>");
}

const OPEN_FENCE = /^\s*(```|~~~)\s*mermaid\s*$/i;
const CLOSE_FENCE = /^\s*(```|~~~)\s*$/;

export interface ShapeCodeBlock {
  source: string;
  /** The text box holds text besides the code block (or more than one block). */
  hasOtherText: boolean;
}

/**
 * Finds the first ```mermaid block in a text box's text. Lines may be paragraphs (\r) or
 * soft line breaks (\v). Returns null when there is no complete block.
 */
export function findShapeCodeBlock(text: string): ShapeCodeBlock | null {
  const lines = text.split(/\r\n|\r|\n|\u000B/).map(normalizeTypedText);
  const start = lines.findIndex((l) => OPEN_FENCE.test(l));
  if (start === -1) return null;
  const fence = lines[start].match(OPEN_FENCE)![1];
  let end = -1;
  for (let i = start + 1; i < lines.length; i++) {
    const close = lines[i].match(CLOSE_FENCE);
    if (close && close[1] === fence) {
      end = i;
      break;
    }
  }
  if (end === -1) return null;
  const outside = [...lines.slice(0, start), ...lines.slice(end + 1)];
  return {
    source: lines.slice(start + 1, end).join("\n"),
    hasOtherText: outside.some((l) => l.trim()),
  };
}

// Mermaid keywords that start a line. PowerPoint capitalizes the first letter of each
// paragraph ("Flowchart TD", "Participant Alice"), which Mermaid doesn't accept.
const LINE_KEYWORDS = new Set(
  (
    "flowchart graph sequenceDiagram classDiagram classDiagram-v2 stateDiagram stateDiagram-v2 erDiagram " +
    "journey gantt pie quadrantChart requirementDiagram gitGraph mindmap timeline zenuml sankey sankey-beta " +
    "xychart xychart-beta block block-beta packet packet-beta kanban architecture-beta radar-beta treemap-beta " +
    "venn-beta ishikawa-beta wardley-beta cynefin-beta treeView-beta swimlane-beta usecase-beta eventmodeling " +
    "participant actor note loop alt else opt par and critical break rect end autonumber activate deactivate " +
    "box create destroy link links properties details title section subgraph direction style classDef class " +
    "linkStyle click callback state commit branch checkout switch merge cherry-pick dateFormat axisFormat " +
    "tickInterval excludes includes todayMarker weekday x-axis y-axis quadrant-1 quadrant-2 quadrant-3 " +
    "quadrant-4 bar line columns space service group junction set union anchor component evolve element " +
    "requirement functionalRequirement interfaceRequirement performanceRequirement physicalRequirement " +
    "designConstraint id text risk verifymethod type docref curve axis max min ticks showLegend graticule " +
    "accTitle accDescr root systemBoundary tf complex complicated clear chaotic"
  ).split(" ")
);

const LINE_START_WORD = /^(\s*)([A-Z])([\w-]*)/;

const escapeRegExp = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Lowercases a line's first word if that gives a Mermaid keyword, unless the word also appears
 * elsewhere in the source: then it is likely a name the user chose (a node called "Text" used as
 * "Text --> B" and "A --> Text"), and changing one occurrence would split it in two.
 */
function lowerFirstKeyword(line: string, source: string): string {
  return line.replace(LINE_START_WORD, (whole, indent: string, first: string, rest: string) => {
    const word = first.toLowerCase() + rest;
    if (!LINE_KEYWORDS.has(word)) return whole;
    const original = first + rest;
    const everywhere = source.match(new RegExp(`(?<![\\w-])${escapeRegExp(original)}(?![\\w-])`, "g"))?.length ?? 0;
    const atLineStart = source.split("\n").filter((l) => l.match(LINE_START_WORD)?.slice(2, 4).join("") === original).length;
    return everywhere > atLineStart ? whole : indent + word;
  });
}

/**
 * Variants of the source with PowerPoint's auto-capitalization undone: first only on the
 * diagram-type line, then on every line. Try them only if the original fails to parse.
 */
export function capitalizationRepairs(source: string): string[] {
  const lines = source.split("\n");
  const firstIndex = lines.findIndex((l) => l.trim() && !l.trim().startsWith("%%"));
  if (firstIndex === -1) return [];
  const firstOnly = [...lines];
  firstOnly[firstIndex] = lowerFirstKeyword(lines[firstIndex], source);
  const all = lines.map((l) => lowerFirstKeyword(l, source));
  return [...new Set([firstOnly.join("\n"), all.join("\n")])].filter((s) => s !== source);
}
