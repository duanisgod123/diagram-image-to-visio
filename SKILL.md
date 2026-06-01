---
name: diagram-image-to-visio
description: Convert uploaded diagram images into editable Microsoft Visio `.vsdx` files by reconstructing shapes, text, containers, and connectors, then generating the drawing through PowerShell and Visio COM automation. Use when Codex needs to turn PNG/JPG/JPEG/BMP screenshots or exported images of flowcharts, business process diagrams, framework diagrams, technical route maps, or box-and-arrow architecture diagrams into editable Visio output with preserved core layout, connection logic, solid/dashed line styles, and basic text styling.
---

# Diagram Image to Visio

## Overview

Use native vision reasoning to read the uploaded diagram, normalize it into a structured spec, then call the bundled PowerShell generator to build an editable Visio document through COM automation.

Prioritize editable shapes, correct connector topology, and retained core styling over pixel-perfect visual effects.

## Workflow

1. Read the uploaded image directly.
2. Infer whether the image is within scope:
   - Good fit: flowcharts, process maps, framework diagrams, box-and-arrow technical diagrams, research route maps, layered architecture diagrams.
   - Weak fit: hand-drawn sketches, dense scientific plots, irregular illustrations, screenshots with heavy gradients or decorative textures.
3. Extract a normalized diagram spec in the schema described in [references/diagram-spec.md](references/diagram-spec.md).
4. For non-trivial diagrams, run [scripts/Analyze-DiagramImage.ps1](scripts/Analyze-DiagramImage.ps1) to get canvas size, dominant colors, and coarse visual components before authoring the spec.
5. If `D:\Desktop\visio模板库` or another template library exists, run [scripts/Build-VisioTemplateIndex.ps1](scripts/Build-VisioTemplateIndex.ps1) and inspect the previews/index before deciding whether to redraw or reuse a style/module.
6. Run [scripts/Test-VisioEnvironment.ps1](scripts/Test-VisioEnvironment.ps1) before first use in a session or after any COM failure.
7. Save the spec JSON in the current workspace or another writable temp location.
8. Prefer [scripts/Convert-DiagramImageToVisio.ps1](scripts/Convert-DiagramImageToVisio.ps1) for normal runs. It accepts the source image, defaults output to the same basename with `.vsdx`, probes Visio, and can clean the generated JSON with `-CleanupIntermediate`.
9. Use [scripts/Convert-DiagramSpecToVisio.ps1](scripts/Convert-DiagramSpecToVisio.ps1) directly only when debugging the generator or when no source image path is involved.
10. Export the generated `.vsdx` back to PNG with [scripts/Export-VisioPreview.ps1](scripts/Export-VisioPreview.ps1), then compare it with the source image. Run [scripts/Compare-DiagramPreview.ps1](scripts/Compare-DiagramPreview.ps1) to produce a JSON diff and heatmap. Do not stop after the first successful file write when visible differences remain.
11. Run at least five source-vs-preview comparison and revision rounds for every non-trivial diagram. Each round must export a fresh preview, compare it against the source, make targeted fixes, and record what changed. Only skip the five-round minimum for trivial diagrams where the first preview is already visually correct, and state that exception explicitly.

## Extraction Rules

Build the spec conservatively.

- Represent every meaningful node as an individual editable shape.
- Preserve text exactly when legible. Do not paraphrase labels.
- Preserve solid vs dashed vs dotted line intent.
- Preserve base fill color, line color, text color, and approximate font size when visible.
- Preserve bold vs regular text when it materially affects hierarchy.
- Normalize gradients, shadows, glow, bevels, and decorative effects to flat fills and plain outlines.
- Mark uncertain edges with `uncertain: true` instead of guessing silently.
- Prefer correct `from` and `to` relationships over decorative intermediate bends.
- Treat big dashed or framed regions as `containers` when they semantically group child nodes.
- Treat legends, footnotes, and corner badges as `annotations` if they are useful but not part of the core flow.
- For large section-to-section arrows, use native arrow shapes in spec, not text glyph placeholders.
- Use image crops only when the content is too complex to reconstruct cleanly, such as point clouds, voxel cubes, dense formulas, or detailed icons. Never paste a large diagram chunk merely to save time when it can reasonably be rebuilt from editable rectangles, text, and connectors.
- Prefer editable reconstruction for repeated blocks: model vertical stacks as a frame plus editable child rectangles; model legends as an editable frame plus labels, swatches, and line samples; model tables or metric panels as grouped editable boxes when text is legible.
- If a raster crop is used for visual fidelity, consider adding editable text/shapes for the same semantic module when the user is likely to revise labels later.
- When sibling boxes clearly follow the same grid, express that with `layout` constraints instead of hand-tuning every coordinate.
- When icons repeatedly dock to box edges, use `slot` instead of repeating absolute icon coordinates.
- When connector bend geometry matters, use `render_mode: "polyline"` with explicit `waypoints`.
- Default imported icon crops to auto-trim white margins.
- For screenshot fragments with multiple isolated regions, use `trim_mode: "largest-component"` so generator keeps only main visual payload.
- Default box text to auto-fit downward so text stays inside its own shape before other fidelity tweaks.

When diagram is crowded, first recover:

1. Top-level sections or containers
2. Main nodes
3. Main connectors
4. Secondary annotations and legend items

## Iterative Visual Comparison

For every medium or complex diagram, use this acceptance loop before final response. Complete at least five comparison-and-revision rounds, even if the file writes successfully earlier:

1. Generate or update the `.vsdx`.
2. Export a fresh PNG preview from Visio.
3. Run `Compare-DiagramPreview.ps1` when a source bitmap is available, then inspect the heatmap hot regions. Also compare source vs preview region by region: title/header, main flow, side panels, legends, lower sections, arrows, and dense modules.
4. Fix obvious differences yourself, especially clipped text, wrong crop bounds, misplaced icons, missing arrows, uneven spacing, wrong module sizes, poor alignment, or modules that should be editable but were pasted as images.
5. Repeat steps 1-4 until at least five rounds are complete and no large visual or layout issue remains. If stopping with known approximations after five rounds, state them plainly.

Do not give a "file generated" final answer after the first successful run for any non-trivial diagram. The default expectation is five rounds of source-vs-preview comparison and revision before final delivery.

## Editability Priority

High-fidelity output should still be editable where feasible.

- Reconstruct text-bearing boxes, legends, simple charts, process stacks, labels, arrows, dashed regions, and section bands as native Visio shapes/text.
- Use raster crops for hard visual content: point clouds, 3D voxel drawings, detailed scientific mini-illustrations, dense equations, and logos/icons that would take longer to redraw than they are worth.
- For a mixed module, use a native outer frame and editable labels, then crop only the complex icon/graphic inside it.
- For formulas and metric panels, split the module into native title text, formula text, explanatory note text, and only the truly hard visual fragments. Do not crop the whole metric panel when text is legible.
- Use the original image as a positioning underlay during iteration: align native boxes and crops to source coordinates, export preview, compare, and adjust. Do not leave a full-page source screenshot in the final Visio file.
- If Visio can draw a part cleanly, redraw it. If Visio drawing would be visibly worse, use a tight local screenshot for that subpart only.
- If a native redraw is semantically useful but visually unstable, use a hybrid stack: keep native invisible anchors/text placeholders for connector topology, place a tight local crop as the visible layer, and verify no draft native layer leaks into the preview.
- When replacing a crop with editable shapes, verify the new native version does not introduce wrapped text, clipped text, overlap, or uneven white space. If native text becomes unreadable, shorten line breaks, enlarge the host box, reduce font size, or fall back to a hybrid crop plus editable overlay.
- Keep connectors attached to native nodes or invisible anchors, not to visual-only crops, when the relationship will be edited later.
- Match source typography: use common Office fonts such as Arial/Calibri for labels and Cambria Math for formulas; tune font size after preview export. Very small source labels may need 4-5 pt text.

## Template Library

Before manually recreating complex scientific or presentation-style modules, check the user's local template library when available:

`D:\Desktop\visio模板库`

Use matching or easily adapted `.vsdx` templates as references for style, spacing, icon vocabulary, or reusable Visio groups. Do not blindly copy unrelated templates; pick only assets that improve fidelity or speed without harming editability.

## Layout Audit

Before final delivery, inspect the exported preview for:

- overlapping text, icons, connectors, or frames
- clipped labels or crops
- large unused blank areas caused by misplaced groups
- uneven spacing between parallel modules
- arrows crossing through important text or icons
- rows/columns that should align but drift
- text that auto-wraps into vertical fragments
- hidden or draft reconstruction layers accidentally visible underneath a local crop

Fix these issues proactively. If a dense original image forces a trade-off, prefer no overlap and readable modules over exact-but-broken placement.

## Default Output Behavior

If user gives only an image:

- Save output beside the source image with the same basename and `.vsdx` extension.
- Use `Orientation=Auto`.
- Use `PageSize=Auto`.
- Keep color unless user asks otherwise.
- Use `#FF00FF` for uncertain connectors.

If user passes options such as page size, orientation, or color preservation, pass them through to the PowerShell script.

## File Naming and Cleanup

- Name the final `.vsdx` file exactly after the source image basename. Example: `1.png` -> `1.vsdx`.
- During normal runs, delete generated intermediate artifacts after the final `.vsdx` is confirmed written. This includes temporary JSON specs, cropped icon images, trace files, and other scratch outputs.
- Keep intermediate artifacts only when the user is explicitly testing, debugging, or asks to inspect them.

## Generator Contract

The bundled PowerShell generator expects:

- A JSON spec matching [references/diagram-spec.md](references/diagram-spec.md)
- Image-space coordinates with origin at top-left
- Node coordinates expressed as center points
- Connector endpoints expressed by node ids plus side hints, with optional exact normalized glue points

Do not write ad hoc PowerShell each time unless the bundled script is insufficient. Prefer updating the script if you discover a repeated need.

## Regression Lessons From `D:\Desktop\1.png`

These fixes are now part of the expected workflow when tuning complex PPT-style diagrams:

- Always verify the source bitmap's real dimensions before authoring or tuning the spec. A stale `canvas.height` can compress the footer, shift vertical spacing, and make an otherwise correct Visio page look wrong.
- For header and footer decorative plates, crop from the true source bounds and keep `trim_white_margin: false`; auto-trimming can make corner art too small or move it away from the edge.
- Keep large decorative footer/header images as raster fragments, but keep all core boxes, titles, labels, and connectors as editable Visio shapes/text.
- When side note boxes contain both a large icon and small text, model them as an empty rectangle, an imported icon crop, and a separate editable text annotation. This avoids Visio text auto-fit crushing the label or colliding with the icon.
- For icon-bearing process boxes, reserve text padding on the icon side. Do not let centered text overlap a left-docked icon.
- Recreate repeated rows from source coordinates when spacing is visibly non-uniform; do not force a `distribute` rule if the original diagram intentionally has varied gaps.
- For explicit polyline branch connectors, use Visio-style glue coordinates: `y=1` means top and `y=0` means bottom. Wrong glue-space interpretation flips endpoints and routes branches through boxes.
- After each manual tuning pass, export the generated `.vsdx` back to PNG and inspect the exact circled problem areas before overwriting the user's target file.
- When a user marks bad icons, separate two failure modes:
  - **Crop failure**: icon is incomplete or includes unrelated border/blank area. Re-crop from the original source bitmap, using a zoomed coordinate grid if needed.
  - **Placement failure**: icon crop is complete but appears in the wrong slot. Keep the crop and adjust `x`, `y`, `width`, `height`, and host node `text_padding`.
- Do not place icons by raw original-image coordinates when the reconstructed box is narrower or shifted from the source. Place the icon relative to the reconstructed host box's left icon slot, then reserve enough left text padding so labels do not overlap.
- For repeated icon rows in another column, make them visually match the best-correct column: complete icon crop, same slot alignment, no border occlusion, and no text collision.
- If a crop contains frame lines from its host box, crop tighter or shift the source crop inward. Do not mask the problem by shrinking the rendered icon until it becomes illegible.

## Regression Lessons From `2.png` DCR-PCGC Diagram

- Do not leave dense scientific diagrams as mostly raster. Convert easy modules first: legends, repeated convolution stacks, section bars, dashed containers, labels, and connector logic.
- A repeated vertical network stack should be a native outer frame plus editable repeated rectangles. Test the exported preview because Visio text wrapping can turn narrow labels into vertical fragments.
- A legend should usually be native: outer rounded rectangle, editable title/text, color swatches, tensor/voxel samples, and line samples. Use shorter editable labels if exact wording creates overlap.
- Small source text needs true small font sizes. The generator must not clamp all text to about 5.8 pt, or editable labels will wrap and overlap.
- Very dense scientific figures can need 3-5 pt Visio text. The generator's fit logic must allow this range and should support per-item `min_font_size`.
- Metric/loss panels such as "Rate Loss" should be decomposed into separate editable text boxes for the title, formula, and notes instead of pasted as a single image.
- Latent, quantization, hyperprior, and entropy coding boxes should be native frames plus native labels and only cropped inner icons/mini charts when feasible.
- For complex 3D point-cloud/voxel miniatures and formulas, raster crops are acceptable, but crop at the module boundary and keep native anchors/connectors around them.
- When replacing a raster crop with native shapes, compare again; improved editability can still make visual layout worse if text boxes wrap or spacing becomes uneven.
- If the user marks "this could be editable", treat it as a rule for future similar modules, not only the one marked instance.
- When Visio distorts or misplaces a tiny repeated stack after connector gluing or text fitting, stop fighting the native version after one correction pass. Use an exact local crop for the visible layer, hide the native draft layer, and leave an invisible host node for incoming/outgoing connectors.
- Do not rely on visual memory after editing the spec. Export a fresh preview after every material font, crop, z-index, or connector change.

## Failure Handling

If recognition quality is weak:

- Still produce a best-effort spec when the main structure is recoverable.
- Color uncertain connectors with the special uncertain color.
- Tell the user which parts were approximated: unreadable text, ambiguous arrow ownership, overlapping routes, unsupported decorative shapes.

If Visio COM fails:

1. Run `Test-VisioEnvironment.ps1`
2. Check that desktop Visio is installed and licensed
3. Retry with `-Visible $true` if debugging needs an interactive window
4. Report the exact COM error

If a "waiting for printer connection" dialog appears, do not use any print command. The conversion path should use `Page.Export`, not `PrintOut`; the dialog usually comes from Visio querying the default printer while opening or laying out a document. Keep `AlertResponse = 7`, cancel the dialog if it blocks automation, then retry export.

## Resources

- [references/diagram-spec.md](references/diagram-spec.md): JSON schema and authoring rules for the intermediate diagram spec
- [references/visio-generation-notes.md](references/visio-generation-notes.md): Shape mapping, routing choices, and approximation policy
- [scripts/Test-VisioEnvironment.ps1](scripts/Test-VisioEnvironment.ps1): COM environment probe
- [scripts/Analyze-DiagramImage.ps1](scripts/Analyze-DiagramImage.ps1): coarse source analysis for canvas, colors, components, and likely module regions
- [scripts/Build-VisioTemplateIndex.ps1](scripts/Build-VisioTemplateIndex.ps1): export/index local `.vsdx` template libraries for visual/style reuse
- [scripts/Export-VisioPreview.ps1](scripts/Export-VisioPreview.ps1): export a generated Visio page to PNG for preview/diff without ad hoc COM snippets
- [scripts/Compare-DiagramPreview.ps1](scripts/Compare-DiagramPreview.ps1): compare source and exported preview, producing JSON hot regions and optional heatmap
- [scripts/Convert-DiagramImageToVisio.ps1](scripts/Convert-DiagramImageToVisio.ps1): Image/spec wrapper with default output naming, Visio probe, and optional intermediate cleanup
- [scripts/Convert-DiagramSpecToVisio.ps1](scripts/Convert-DiagramSpecToVisio.ps1): Main Visio generator
