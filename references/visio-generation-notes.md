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
