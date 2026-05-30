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
4. Run [scripts/Test-VisioEnvironment.ps1](scripts/Test-VisioEnvironment.ps1) before first use in a session or after any COM failure.
5. Save the spec JSON in the current workspace or another writable temp location.
6. Prefer [scripts/Convert-DiagramImageToVisio.ps1](scripts/Convert-DiagramImageToVisio.ps1) for normal runs. It accepts the source image, defaults output to the same basename with `.vsdx`, probes Visio, and can clean the generated JSON with `-CleanupIntermediate`.
7. Use [scripts/Convert-DiagramSpecToVisio.ps1](scripts/Convert-DiagramSpecToVisio.ps1) directly only when debugging the generator or when no source image path is involved.
8. Verify that the output file exists. If possible, inspect the generated Visio file for obvious routing or text failures.

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
- Use image crops only for small missing icons or decorative corner assets. Never paste large diagram chunks as fallback.
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

## Resources

- [references/diagram-spec.md](references/diagram-spec.md): JSON schema and authoring rules for the intermediate diagram spec
- [references/visio-generation-notes.md](references/visio-generation-notes.md): Shape mapping, routing choices, and approximation policy
- [scripts/Test-VisioEnvironment.ps1](scripts/Test-VisioEnvironment.ps1): COM environment probe
- [scripts/Convert-DiagramImageToVisio.ps1](scripts/Convert-DiagramImageToVisio.ps1): Image/spec wrapper with default output naming, Visio probe, and optional intermediate cleanup
- [scripts/Convert-DiagramSpecToVisio.ps1](scripts/Convert-DiagramSpecToVisio.ps1): Main Visio generator
