# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Office.js task-pane add-in for PowerPoint that renders Mermaid (pinned `mermaid@12.2.0`) and inserts diagrams as SVG pictures. It is a port of Mermaid2Word (https://github.com/tstickland/Mermaid2Word): `src/render/` and `src/templates.ts` are shared code copied from it, so keep fixes there in sync. See README.md for user-facing setup and hosting.

## Commands

- `npm install --ignore-scripts`: a postinstall in `@microsoft/m365agentstoolkit-cli` (pulled in by `office-addin-debugging`) runs `npm config get prefix`, which recent npm rejects. No dependency needs install scripts. CI uses `npm ci --ignore-scripts`.
- `npm start` / `npm stop`: dev server on https://localhost:3443 plus sideload into PowerPoint desktop. The first run installs a localhost dev certificate, which prompts the user.
- `ADDIN_BASE_URL=https://host/path/ npm run build`: production `dist/`. Webpack rewrites `manifest.xml` for production: localhost URLs and the `AppDomain` origin become the hosted URL, the add-in `<Id>` becomes `PROD_ID`, and " (Dev)" is dropped from labels. It also fills the `__ADDIN_BASE_URL__`/`__ADDIN_ID__` placeholders in `install/` (the teammate install page and scripts). Validate with `npx office-addin-manifest validate dist/manifest.xml`. Pushing to `main` does this automatically and publishes to GitHub Pages (`.github/workflows/deploy-pages.yml`).
- `npm run validate`: validates the dev manifest.
- `npx tsc --noEmit -p .`: type-check.
- `npm run test:render`: builds the `--env harness` bundle into `test-build/`. It then renders every entry in `src/templates.ts` in headless **Edge** by default (`BROWSER=chrome`/`firefox` use Playwright's builds, installed with `npx playwright install chromium firefox`, output in `test-output/<browser>/`). Writes SVG/PNG/`.meta.txt` and `placements.json` (where Insert would put each diagram on an empty 16:9 slide) to `test-output/`. It also runs `__selfTest()` (metadata round trip incl. Word alt text, code-block parsing, auto-capitalization repair, layout math). `RAW=id1,id2` also saves Mermaid's unsanitized output.
- `npm run test:powerpoint` (Windows + PowerPoint installed): `scripts/powerpoint-com-test.ps1` adds each SVG to a windowless presentation via COM `Shapes.AddPicture` at its `placements.json` box, tags it like the add-in does, saves `all-diagrams.pptx`/`.pdf` and `slides/<id>.png`, then reopens the file and checks tags and alt text. Look at `test-output/slides/*.png` to see PowerPoint's actual rendering. Run `test:render` first.

The COM test checks rendering and storage only. The Office.js layer (`src/powerpoint/diagrams.ts`) can only be exercised in PowerPoint via `npm start`.

## Architecture

Render pipeline: `mermaidRenderer.ts` → `svgSanitize.ts` (→ `rasterize.ts` for PNG-only) → `powerpoint/diagrams.ts`. The task pane (`taskpane.ts`) only wires UI to `powerpoint/diagrams.ts`.

- **Rendering happens in an off-screen stage div**, `#mermaid-stage`. It must stay laid out: do **not** use `visibility:hidden` or `display:none`, because the sanitizer measures text with `getBoundingClientRect`. Mermaid calls are serialized through a promise queue because Mermaid holds global config.
- **`svgSanitize.ts` works around Office's limited SVG renderer** (same renderer as Word): CSS resolved to presentation attributes, `<foreignObject>` converted to `<text>`, `dominant-baseline` emulated, filters removed (Office drops filtered elements), newlines in `<text>` normalized.
- **Inserting pictures**: PowerPointApi has no add-picture method, so `placeDiagram` uses the Common API `setSelectedDataAsync` with `CoercionType.XmlSvg` (PowerPoint makes its own PNG fallback), or `Image` for PNG-only. It inserts on the *current* slide, so `placeDiagram` first selects the target slide (`setSelectedSlides`) and clears the shape selection. The new shape is found by diffing the slide's shape IDs before and after (retried briefly). `imageLeft`/`imageTop` are always passed: without them PowerPoint would put the picture into a placeholder.
- **Source of truth is the shape tag `MERMAID_DIAGRAM`** (PowerPoint upper-cases tag keys), in the same format as Mermaid2Word's alt text: `%% mermaid-word {json options + w/h natural px + base64 "src"}` followed by the raw source (`sourceMeta.ts`). With PowerPointApi 1.10 the same text also goes in the alt text, and is read back from alt text when there is no tag. That is how pictures pasted from Word are recognized. The marker stays `mermaid-word` for cross-app compatibility.
- **Update/re-render replace the shape**: insert the new picture, delete the old one, then send the new one backward (`setZOrder`, 1.8) once per shape that was above the old one. `layout.rescale` keeps the user's scale relative to the natural size (`w`/`h` in the metadata) and the old center, and keeps a picture that was on the slide on the slide.
- **Insert placement** (`freeArea`): an empty Body/Content placeholder is used and then deleted. Otherwise the area is below a Title placeholder in the top half, or the slide minus a 5% margin. Placeholder types need 1.8, and the slide size needs 1.10 (else 960×540 pt is assumed).
- **Convert code blocks**: a text box (TextBox/Placeholder/GeometricShape) whose text holds a ```` ```mermaid ```` block. Text boxes with other text are reported, not converted. `normalizeTypedText` maps AutoCorrect characters (smart quotes, dashes, `→ ← ⇒ ⇐ ⇔`). If Mermaid rejects the source, `capitalizationRepairs` retries with PowerPoint's auto-capitalized first words lowercased: first the diagram-type line only, then every line, and only for known keywords. The diagram is fitted into the text box's bounds.
- Requirement sets: 1.5 is required (checked in `Office.onReady`); 1.8 and 1.10 features are used when `isSetSupported` says so, each guarded with try/catch so a failure there never blocks the insert.
- ZenUML is registered via `registerExternalDiagrams`. The `logos` icon pack is bundled and lazy-loaded; other Iconify sets are fetched per icon (`iconPacks.ts`), so `test:render` needs internet for `architecture-icons`.

## Gotchas

- Team distribution is by per-user sideload. The Windows installer writes `HKCU\Software\Microsoft\Office\16.0\Wef\Developer\<PROD_ID>` = manifest path, the same mechanism `npm start` uses. The Mac installer copies the manifest into PowerPoint's `wef` folder (`com.microsoft.Powerpoint` container). Changing `PROD_ID` would orphan existing installs. Both IDs differ from the Word add-in's.
- `typescript` is pinned to 5.x: ts-loader cannot use TypeScript 7.
- The task pane stylesheet is copied as a static file (`taskpane.css`) and linked from `taskpane.html`, not imported.
- PowerPoint COM has a single instance: `New-Object -ComObject PowerPoint.Application` attaches to a PowerPoint the user has open. `powerpoint-com-test.ps1` only quits or stops PowerPoint if it wasn't running before. `Quit()` can leave a windowless `POWERPNT` behind, which the script then stops. Never stop all POWERPNT processes.
