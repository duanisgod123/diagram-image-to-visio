# Diagram Spec

Use this schema as intermediate representation between image understanding and Visio generation.

## Goals

- Keep structure explicit.
- Keep coordinates image-relative.
- Preserve core visual hierarchy, not only semantics.
- Make uncertainty visible instead of implicit.

## Coordinate System

- `canvas.width` and `canvas.height` use source image units, usually pixels.
- Origin is top-left.
- `x` and `y` on shapes are center coordinates in source units.
- `width` and `height` use source units.

PowerShell generator rescales this coordinate system to selected Visio page.

## Top-Level Schema

```json
{
  "title": "Example Diagram",
  "canvas": { "width": 1880, "height": 1158 },
  "page": {
    "orientation": "Landscape",
    "page_size": "Auto",
    "background": "#FFFFFF",
    "margin": 0.15
  },
  "defaults": {
    "font_family": "Microsoft YaHei",
    "font_size": 12,
    "font_size_mode": "fit",
    "font_bold": false,
    "text_color": "#111111",
    "line_color": "#333333",
    "line_pattern": "solid",
    "fill_color": "#FFFFFF",
    "vertical_align": "middle"
  },
  "nodes": [],
  "containers": [],
  "connectors": [],
  "annotations": [],
  "images": [],
  "layout": []
}
```

`page.margin` is page-space margin in inches after scaling. Use smaller values when source diagram already has tight outer framing.

## Nodes

Use `nodes` for editable diagram objects that should move independently.

```json
{
  "id": "node-order-processing",
  "shape": "rounded-rectangle",
  "text": "订单处理",
  "x": 972,
  "y": 366,
  "width": 164,
  "height": 68,
  "fill_color": "#D9F2A7",
  "line_color": "#333333",
  "text_color": "#111111",
  "line_pattern": "solid",
  "line_weight": 1.5,
  "font_family": "Microsoft YaHei",
  "font_size": 14,
  "font_size_mode": "fit",
  "font_bold": true,
  "vertical_align": "middle",
  "text_padding": { "left": 8, "right": 8, "top": 4, "bottom": 4 },
  "rotation": 0,
  "z_index": 30
}
```

### Supported `shape` values

- `rectangle`
- `rounded-rectangle`
- `diamond`
- `ellipse`
- `circle`
- `terminator`
- `block-arrow-right`
- `block-arrow-down`

If source shape is unusual, map to closest supported value and note approximation in user-facing summary.

## Containers

Use `containers` for region frames that group child nodes.

```json
{
  "id": "container-logistics-center",
  "text": "物流仓储中心",
  "shape": "rectangle",
  "x": 937,
  "y": 711,
  "width": 1620,
  "height": 942,
  "fill_color": "#FFFFFF",
  "line_color": "#0A7C2F",
  "text_color": "#111111",
  "line_pattern": "dashed",
  "line_weight": 2,
  "font_size": 16,
  "z_index": 1
}
```

Use containers for large dashed or framed areas, not ordinary nodes.

## Connectors

Use `connectors` for logical flow edges.

```json
{
  "id": "edge-inbound-plan",
  "from": "node-inventory",
  "to": "node-shipping",
  "render_mode": "dynamic",
  "from_side": "right",
  "to_side": "left",
  "from_glue": [1.0, 0.5],
  "to_glue": [0.0, 0.5],
  "text": "发货计划",
  "line_color": "#A000B5",
  "text_color": "#222222",
  "line_pattern": "dashed",
  "line_weight": 2.5,
  "arrow_end": "triangle",
  "arrow_size": "small",
  "uncertain": false
}
```

### Connector fields

- `from`, `to`: required node or container ids
- `render_mode`: `dynamic` or `polyline`
- `from_side`, `to_side`: `top`, `right`, `bottom`, `left`, or `auto`
- `from_glue`, `to_glue`: optional normalized Visio-style glue coordinates `[x, y]` inside shape local space, where `x=0` is left, `x=1` is right, `y=0` is bottom, and `y=1` is top
- `waypoints`: optional intermediate source-space points for `polyline` connectors, for example `[[420, 180], [420, 220]]`
- `text`: optional edge label
- `line_pattern`: `solid`, `dashed`, or `dotted`
- `arrow_end`: `none`, `triangle`, or `open`
- `arrow_size`: `tiny`, `small`, `medium`, or `large`
- `uncertain`: boolean

Use glue overrides when branch points need more precision than side midpoints. Keep dynamic connectors for semantic flow; do not encode decorative bend traces unless they materially change readability.

## Annotations

Use `annotations` for titles, notes, legends, and labels that should remain editable but are not core nodes.

```json
{
  "id": "annotation-legend-title",
  "text": "信息流",
  "x": 1610,
  "y": 1002,
  "width": 120,
  "height": 24,
  "text_color": "#111111",
  "font_size": 12,
  "shape": "none",
  "z_index": 100
}
```

If annotation includes small sample line or color swatch, model it as annotation plus optional node or connector sample.

## Layout

Use `layout` for reusable spatial constraints so spec carries layout intent, not only raw coordinates.

```json
[
  {
    "type": "align",
    "ids": ["node-a", "node-b", "node-c"],
    "axis": "x",
    "anchor": "center",
    "reference_id": "node-a"
  },
  {
    "type": "same_size",
    "ids": ["node-a", "node-b", "node-c"],
    "dimension": "width",
    "reference_id": "node-a"
  },
  {
    "type": "distribute",
    "ids": ["node-a", "node-b", "node-c"],
    "axis": "y",
    "anchor": "center",
    "start": 120,
    "gap": 46
  },
  {
    "type": "offset",
    "id": "icon-a",
    "reference_id": "node-a",
    "dx": -74,
    "dy": 0
  }
]
```

### Layout rule types

- `align`: force multiple items to share one anchor on one axis
- `same_size`: force width, height, or both to match reference or explicit value
- `distribute`: place ordered items at equal intervals from `start` plus `gap`, or between `start` and `end`
- `offset`: position one item relative to another item center using source-space offsets

### Anchors

- For `axis: "x"`: `left`, `center`, `right`
- For `axis: "y"`: `top`, `center`, `bottom`

## Images

Use `images` only for small visual fragments that cannot be reconstructed with native Visio primitives, usually icons or decorative corner badges.

```json
{
  "id": "image-lock-badge",
  "path": "C:\\temp\\lock.png",
  "x": 1430,
  "y": 86,
  "width": 96,
  "height": 38,
  "slot": {
    "reference_id": "node-target",
    "side": "left",
    "align": "center",
    "dx": 12,
    "dy": 0
  },
  "trim_white_margin": true,
  "trim_mode": "largest-component",
  "trim_padding": 0,
  "z_index": 60
}
```

`slot` is optional. It places current item relative to host item bounds. Supported `side`: `left`, `right`, `top`, `bottom`, `center`. Supported `align`: `start`, `center`, `end`.

Image fallback rules:

- Crop icon-sized assets only. Do not paste large slices of original diagram as shortcut.
- Trim obvious white margins before import. `trim_white_margin` defaults to `true`.
- If crop contains separated decorative fragments, `trim_mode: "largest-component"` keeps only biggest non-white connected region.
- `trim_padding` optionally keeps extra pixels around trimmed content. Default is `0` for tight crops.
- Keep imported images above host shape but below unrelated text unless source clearly overlays text.
- When native Visio shape exists, prefer native shape over image crop.

## Authoring Rules

1. Keep ids stable and descriptive.
2. Preserve text exactly when readable.
3. Use `uncertain: true` when ownership of arrow or branch is ambiguous.
4. Prefer fewer, more reliable connectors over many speculative connectors.
5. Keep `z_index` low for containers, medium for nodes, high for annotations and top-layer icon crops.
6. Use `font_bold`, `vertical_align`, and `text_padding` when typography materially affects fidelity.
7. Prefer `font_size_mode: "fit"` for box text so long labels shrink before they overflow or collide.
8. Use `block-arrow-right` and `block-arrow-down` for large inter-section arrows instead of text glyph stand-ins.
9. Use `from_glue` and `to_glue` when branch points need more precise anchor placement than side midpoints provide.
10. Prefer `layout` constraints over repeated hand-tuned coordinates when groups should align, match size, or distribute evenly.
11. Use `slot` when icons or helper labels should stay attached to a host box with stable side padding.
12. Use `render_mode: "polyline"` plus `waypoints` only when connector geometry materially affects readability or fidelity.
13. For dense Chinese diagrams, default body text to `Microsoft YaHei` and let generator auto-fit size downward rather than forcing one fixed point size everywhere.
