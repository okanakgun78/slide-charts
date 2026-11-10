# Slide Charts

Consulting-style charts for PowerPoint (stacked, waterfall, Mekko, Gantt, …), drawn as native
PowerPoint shapes. Built as an Office web add-in: works in PowerPoint for Mac, Windows and the web.

**Status:** v0.4 — column and bar charts (stacked, clustered, 100%), line and combination charts,
Gantt/timeline (tasks, milestones, groups; weeks to years), data sheets with paste from Excel,
click-to-edit.

## Privacy

The add-in has no server. Its code is downloaded from GitHub Pages; your chart data is stored only
inside your .pptx file (in hidden shape tags) and never leaves your computer.

## Install on Mac (for testing)

1. In Terminal, run:

   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef && curl -fsSL https://okanakgun78.github.io/slide-charts/manifest.xml -o ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slide-charts.xml
   ```

2. Quit PowerPoint completely (Cmd+Q) and reopen it, then open any presentation.
3. Select **Home › Add-ins** and choose **Slide Charts**. A **Charts** button appears on the Home tab.

Updates are published automatically; reopen the panel to get the latest version.

To remove: delete the file `slide-charts.xml` from the `wef` folder above and restart PowerPoint.

## Install for good (Microsoft 365 admin)

Microsoft 365 admin center › Settings › Integrated apps › Upload custom apps › Office Add-in ›
provide the manifest URL `https://okanakgun78.github.io/slide-charts/manifest.xml` › assign to yourself.

## Development

```bash
npm install
npm test          # engine unit tests
npm run build     # type-check, bundle, write dist/manifest.xml (set BASE_URL for the host)
npm run e2e       # build, then drive the panel against a mock PowerPoint (needs Playwright)
```

Layout: `src/engine` (pure chart layout + SVG preview), `src/ppt` (PowerPoint renderer, diagnostics),
`src/ui` (data sheet, parsing, settings), `src/taskpane.*` (side panel), `test/e2e` (mock PowerPoint +
end-to-end test), `manifest.template.xml` (add-in manifest).
