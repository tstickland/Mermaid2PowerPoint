# Mermaid Diagrams for PowerPoint

A Microsoft PowerPoint add-in that renders [Mermaid](https://mermaid.js.org/) diagrams (Mermaid 12.2)
and inserts them on your slides as crisp, editable pictures. It is the PowerPoint version of
[Mermaid2Word](https://github.com/tstickland/Mermaid2Word).

- **Every diagram type** in the Mermaid docs: Flowchart, Swimlanes, Sequence, Class, State, ER,
  User Journey, Gantt, Pie, Quadrant, Requirement, Use Case, GitGraph, C4, Mindmap, Timeline,
  ZenUML, Sankey, XY Chart, Block, Packet, Kanban, Architecture, Radar, Event Modeling, Treemap,
  Venn, Ishikawa, Wardley, Cynefin and TreeView. The **Template** list has a starter for each.
- **Vector graphics**: diagrams are inserted as SVG, so they stay sharp at any size, on a projector
  and in print. PowerPoint keeps its own PNG fallback for older viewers.
- **Placed for you**: **Insert** puts the diagram on the current slide. It fills an empty content
  placeholder if there is one. Otherwise it fills the area below the title (or the whole slide).
  **Size → Actual size** inserts at natural size instead.
- **Editable**: the Mermaid source is stored with the picture. Click a diagram, then
  **Load selected**, edit, and **Update selected**. The diagram keeps its position, scale and
  stacking order. This also works on diagrams copied to other presentations.
- **Works with Mermaid2Word**: the source is stored in the same format as the Word add-in. A diagram
  copied from Word to PowerPoint (or back) stays editable in the other add-in. Reading a pasted Word
  diagram needs PowerPoint with PowerPointApi 1.10.
- **Convert code blocks**: type a fenced block in a text box and convert it in place. The diagram
  takes the text box's place:

  ````
  ```mermaid
  flowchart LR
    A --> B
  ```
  ````

  PowerPoint's AutoCorrect changes are undone automatically: smart quotes, dashes, `→` for `-->`,
  and capitalized first words such as `Flowchart` or `Participant`.
- **Icon libraries**: Architecture diagrams and flowchart icon shapes can use any of the
  200,000+ icons on [Iconify](https://icon-sets.iconify.design/), written as `prefix:name`:

  ```
  architecture-beta
    service db(mdi:database)[Database]
    service users(fa6-solid:users)[Users]
  ```

  The `logos` set is built in and works offline. Icons from any other set are downloaded from
  the Iconify API (`api.iconify.design`) when the diagram renders. They are then stored in the
  picture, so nobody needs a connection to view the presentation. Offline, Mermaid draws a "?" in
  place of the icon.
- **Options**: theme (default, neutral, dark, forest, base), background, size, and a PNG-only
  mode with its own resolution. **Re-render all** applies the current options to every diagram in
  the presentation, keeping each one's position and scale.

Requires PowerPointApi 1.5: PowerPoint in Microsoft 365 on Windows, Mac and the web. Newer
versions add extras: stacking order is kept on update (1.8), content placeholders are filled (1.8),
the real slide size is used instead of assuming 16:9 (1.10), and the source is also written to alt
text (1.10).

## Use it locally (PowerPoint desktop on Windows/Mac)

```sh
npm install --ignore-scripts
npm start        # trusts a localhost dev certificate (first run asks for confirmation),
                 # starts https://localhost:3443 and opens PowerPoint with the add-in sideloaded
npm stop         # when finished
```

`--ignore-scripts` skips a postinstall script in the Office debugging tools that fails on recent
npm versions. None of the dependencies need install scripts.

In PowerPoint: **Home → Add-ins → Mermaid Diagrams (Dev)** (under Developer Add-ins) opens the task
pane the first time. After that, the **Insert → Mermaid Diagram** button appears.

To sideload by hand instead: run `npm run dev-server`, then in PowerPoint on the web use
**Home → Add-ins → More Add-ins → Upload My Add-in** with `manifest.xml`.

## Share it with your team

Every push to `main` builds the add-in and publishes it to GitHub Pages
(`.github/workflows/deploy-pages.yml`). This needs **Settings → Pages → Source: GitHub Actions**
on the repository. Teammates install from the published page:

**https://tstickland.github.io/Mermaid2PowerPoint/**

It has a one-line installer for PowerPoint on Windows (PowerShell) and Mac (Terminal), plus a
manifest download for PowerPoint on the web. No Microsoft 365 admin access is needed: the
installers sideload the add-in for the current user. Code changes reach everyone automatically on
their next PowerPoint restart. Only changes to `manifest.xml` (ribbon, name, permissions) require
teammates to re-run the installer; bump `<Version>` when you make one.

The shared build has a different add-in ID from the dev manifest (`webpack.config.js` swaps it),
and the dev copy is labelled "(Dev)", so both can be installed side by side. Both IDs differ from
the Word add-in's, so the two add-ins can be installed together. To host somewhere else, build with
`ADDIN_BASE_URL=https://your.host/path/ npm run build` and publish `dist/`.

## Tests

```sh
npm run test:render      # renders all templates in headless Edge through the real pipeline
                         # (BROWSER=chrome or BROWSER=firefox for the PowerPoint-on-the-web engines),
                         # runs self-tests, writes SVG/PNG/metadata to test-output/
npm run test:powerpoint  # (Windows + PowerPoint) inserts every SVG into a real presentation via COM,
                         # saves test-output/all-diagrams.pptx/.pdf and slides/*.png, then reopens
                         # the file and checks the stored source survived
```

## How it works

| File | Purpose |
| --- | --- |
| `src/render/mermaidRenderer.ts` | Configures Mermaid (plain-SVG labels, ZenUML, icon packs) and renders. |
| `src/render/iconPacks.ts` | Finds `prefix:name` icon references and fetches just those icons from the Iconify API. |
| `src/render/svgSanitize.ts` | Makes the SVG Office-safe: resolves CSS into attributes, converts HTML labels to `<text>`, emulates `dominant-baseline`, removes filters. |
| `src/render/rasterize.ts` | Builds the PNG for PNG-only mode. |
| `src/powerpoint/diagrams.ts` | Insert, update, convert-code-blocks and re-render operations using PowerPoint's JavaScript API. |
| `src/powerpoint/layout.ts` | Where diagrams go and how big they are. |
| `src/powerpoint/codeBlocks.ts` | Finds fenced blocks in text boxes and undoes AutoCorrect. |
| `src/powerpoint/sourceMeta.ts` | Encodes the source and options (same format as Mermaid2Word). |
| `src/templates.ts` | Starter templates. |
