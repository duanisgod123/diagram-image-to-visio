# Visio Generation Notes

## Scope

This skill targets diagrams that can be reconstructed as boxes, arrows, containers, imported icon fragments, and dynamic connectors.

Good first-pass sources:

- Process flowcharts
- Business workflow maps
- Research framework diagrams
- Architecture and layered box diagrams
- PPT or Word exported route maps

Less reliable sources:

- Hand-drawn sketches
- Heavy gradients and shadow-rich marketing graphics
- Scientific figures with tiny labels and overlapping arrows
- Diagrams where connectors are mostly freeform curves

## Approximation Policy

Preserve:

- Shape independence
- Text editability
- Core connector relationships
- Solid/dashed/dotted distinction
- Basic fill, line, and text colors
- Approximate font size
- Typography weight when visually important
- Section-to-section block arrows when they carry flow meaning
- Group alignment and regular spacing when source clearly uses them
- Text containment inside boxes before visual similarity tricks

Approximate:

- Gradients
- Shadows
- Glow
- Decorative iconography only when no native Visio primitive exists
- Minor routing bends after main topology is preserved
- Slight connector anchor offsets that do not change meaning

## Shape Mapping

Use these mappings in PowerShell generator:

- `rectangle` -> `Page.DrawRectangle`
- `rounded-rectangle` -> rectangle plus `Rounding`
- `diamond` -> closed polyline
- `ellipse` and `circle` -> `Page.DrawOval`
- `terminator` -> rounded rectangle with heavier rounding
- `block-arrow-right` -> closed polyline arrow
- `block-arrow-down` -> closed polyline arrow
- `none` -> text-only annotation

## Connector Policy

- Use dynamic connectors from Visio connector tool data object.
- Glue begin and end points to specified side hints.
- When spec provides `from_glue` or `to_glue`, prefer those exact normalized anchors over side midpoints.
- Use `render_mode: "polyline"` only for visually important orthogonal branches or explicit bend geometry.
- Default small connector arrows to smaller arrowhead size unless source clearly uses oversized heads.
- If source diagram contains visible label on edge, assign it to connector text.
- If relationship is uncertain, color connector with configured uncertain color instead of silently guessing.

## Layout Policy

- Apply `layout` normalization before mapping source coordinates to Visio page coordinates.
- Use layout constraints to express repeated rows, columns, and spacing patterns instead of hand-tuning every sibling item.
- Keep raw coordinates as fallback seeds; layout rules may refine them.

## Slot Policy

- Resolve `slot` after layout normalization so host geometry is final.
- Use slots for icons, badges, or helper labels that must stay docked to a host box.
- Prefer slots over manual x/y duplication when many items share same left-icon pattern.

## Image Trim Policy

- Imported icon crops should be auto-trimmed for white margins by default.
- Keep only 1-2 pixels of safety padding after trim.
- If trim would cut meaningful stroke pixels, allow opt-out with `trim_white_margin: false`.
- For bad screenshot-like crops that include multiple separated regions, prefer `trim_mode: "largest-component"`.
- When a crop is incomplete, return to the source bitmap and re-crop with a coordinate grid rather than stretching the existing crop.
- When a crop is complete but misplaced, adjust the rendered image position relative to the reconstructed host box; do not re-crop.
- If the host box width differs from the source bitmap, place icon crops by the host box's left slot and increase `text_padding.left` so the editable text starts after the icon.
- If a crop includes host border lines or large blank margins, tighten the crop inward and keep the rendered size close to the source icon's visible size.

## Image Fallback Policy

- Imported image fragments are last-resort support for missing icons, not substitute for whole modules.
- Keep crops tight to icon bounds.
- Respect `z_index`; do not blindly bring every image to front.
- Header and footer decorative plates are acceptable image fragments when they do not interfere with editability of central business content.

## Editable-First Module Policy

- Rebuild simple repeated modules as native shapes before using a bitmap crop. Good candidates: legends, vertical process stacks, metric boxes, section headers, simple tables, and repeated icon-label rows.
- Use hybrid reconstruction for complex modules: native outer box/title/connectors plus cropped inner visual when the inner content is a point cloud, 3D voxel drawing, dense formula, or detailed icon.
- After replacing a crop with native shapes, export a preview and inspect text wrapping. Narrow Visio text boxes can make labels unreadable even when source coordinates are correct.
- Prefer slightly simplified editable labels over exact labels that overlap or wrap vertically.
- Keep invisible anchors or native host nodes for connector endpoints when a visual crop is used, so future editing does not depend on selecting a bitmap.
- Hybrid fallback is valid when native Visio output is worse than a tight crop: keep invisible native anchors/text placeholders for semantics and connectors, render the exact crop on top, and ensure no draft layer is visible in the exported preview.
- For formulas or loss/metric panels, split title, equation, and note into separate native text boxes. Use Cambria Math or a similar math-friendly font for equations.
- For latent/quantization/hyperprior/entropy boxes, draw the frame and labels natively, then insert tight crops only for internal voxel/chart/icon visuals.
- Allow true small text sizes when matching scientific figures. A hard minimum around 5.8-7 pt causes label wrapping and should be avoided; 3-5 pt may be necessary for dense labels.
- Use the source image as a positioning reference throughout iteration; do not leave a full-image underlay in the delivered `.vsdx`.

## Comparison Loop

- A successful `.vsdx` write is not completion for complex diagrams. Export a PNG preview and compare it against the source.
- Run separate checks for visual fidelity and editability: modules that look right but are pasted as avoidable raster chunks should be revisited.
- Fix obvious discrepancies without waiting for the user to circle them: clipped crops, wrong icon positions, missing dashed flow lines, uneven spacing, and text overflow.
- Use available local Visio templates, especially `D:\Desktop\visio模板库`, as style/asset references before hand-building a complex module from scratch.
- Use `Compare-DiagramPreview.ps1` to generate a heatmap and hot-region JSON. Treat hot regions as a work queue, then inspect them visually before making changes.
- Use `Export-VisioPreview.ps1` for repeatable PNG previews instead of handwritten COM snippets. It uses Visio `Page.Export`; this is export, not printing.
- Use `Analyze-DiagramImage.ps1` before spec authoring when no better OCR/layout output is available. Its components are rough visual hints, not final semantic boxes.
- Use `Build-VisioTemplateIndex.ps1` to create a preview/index of local templates; inspect likely matches before deciding style, line weight, or reusable modules.
- If Visio shows a printer connection dialog during automation, cancel it and retry export. The pipeline should never call `PrintOut`; the dialog is usually Visio/default-printer layout probing.

## Recommended User-Facing Summary

After generation, report:

1. Output file path
2. Whether generation completed
3. Count of nodes, containers, connectors, and uncertain connectors
4. Any approximations:
   - unreadable text
   - merged or simplified decorative elements
   - uncertain connector ownership
   - unsupported irregular shapes
