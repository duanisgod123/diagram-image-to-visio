# Diagram Image to Visio

Convert diagram screenshots into editable Microsoft Visio `.vsdx` files.

This repository packages a Codex skill plus PowerShell scripts for reconstructing flowcharts, architecture diagrams, process maps, framework diagrams, and technical route maps as native Visio shapes, connectors, containers, annotations, and image crops where needed.

The priority is editability and correct connector topology, not pixel-perfect visual effects.

## What It Does

- Reads an uploaded diagram image with a vision-capable AI agent.
- Normalizes the diagram into a structured JSON spec.
- Uses PowerShell and Microsoft Visio COM automation to generate an editable `.vsdx`.
- Exports Visio previews to PNG for visual comparison and iterative revision.
- Supports native nodes, containers, connectors, annotations, layout constraints, arrow shapes, and selective raster crops for complex visual fragments.

## Best Fit

- Flowcharts
- Business process maps
- Box-and-arrow architecture diagrams
- Layered technical diagrams
- Research framework diagrams
- PPT or Word exported route maps

Less suitable inputs include hand-drawn sketches, dense plots, highly decorative graphics, or images whose labels and arrow ownership are not recoverable.

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or newer
- Microsoft Visio desktop edition, installed and licensed
- A vision-capable AI coding agent that can read `SKILL.md` and author the intermediate JSON spec

## Repository Layout

```text
diagram-image-to-visio/
|-- README.md
|-- SKILL.md
|-- agents/
|   `-- openai.yaml
|-- references/
|   |-- diagram-spec.md
|   `-- visio-generation-notes.md
`-- scripts/
    |-- Analyze-DiagramImage.ps1
    |-- Build-VisioTemplateIndex.ps1
    |-- Compare-DiagramPreview.ps1
    |-- Convert-DiagramImageToVisio.ps1
    |-- Convert-DiagramSpecToVisio.ps1
    |-- Export-VisioPreview.ps1
    `-- Test-VisioEnvironment.ps1
```

## Quick Start

Clone this repository:

```powershell
git clone https://github.com/duanisgod123/diagram-image-to-visio.git
cd diagram-image-to-visio
```

Check the Visio COM environment:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-VisioEnvironment.ps1
```

Expected successful output is similar to:

```json
{"installed":true,"name":"Microsoft Visio","version":"16.0","message":"Visio COM available."}
```

Give your image to a vision-capable agent and ask it to follow `SKILL.md`. The agent should create a JSON spec that follows `references/diagram-spec.md`, then run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Convert-DiagramImageToVisio.ps1 `
  -ImagePath "D:\path\to\diagram.png" `
  -SpecPath "D:\path\to\diagram.json"
```

By default, output is saved beside the source image with the same basename and a `.vsdx` extension.

## Typical Agent Workflow

1. Read the source image.
2. Extract a normalized diagram spec using `references/diagram-spec.md`.
3. Run `scripts/Test-VisioEnvironment.ps1` if Visio COM has not been checked in the current session.
4. Generate the `.vsdx` with `scripts/Convert-DiagramImageToVisio.ps1`.
5. Export a PNG preview with `scripts/Export-VisioPreview.ps1`.
6. Compare source and preview with `scripts/Compare-DiagramPreview.ps1`.
7. Revise the JSON spec and repeat preview comparison until layout, text, and connector quality are acceptable.

For medium or complex diagrams, the skill expects at least five source-vs-preview comparison and revision rounds unless the first result is already visually correct for a trivial image.

## Intermediate Spec

The generator consumes a JSON file with image-space coordinates and top-left origin. A minimal example:

```json
{
  "title": "Example",
  "canvas": { "width": 1200, "height": 800 },
  "page": { "orientation": "Landscape", "page_size": "Auto" },
  "nodes": [
    {
      "id": "start",
      "shape": "rounded-rectangle",
      "text": "Start",
      "x": 200,
      "y": 120,
      "width": 140,
      "height": 60,
      "fill_color": "#EAF4FF",
      "line_color": "#2F5597"
    }
  ],
  "connectors": []
}
```

See `references/diagram-spec.md` for the full schema and authoring rules.

## Main Scripts

- `scripts/Analyze-DiagramImage.ps1`: reads source dimensions, colors, and coarse component regions.
- `scripts/Test-VisioEnvironment.ps1`: checks whether Visio COM automation is available.
- `scripts/Convert-DiagramImageToVisio.ps1`: normal conversion wrapper for image plus JSON spec.
- `scripts/Convert-DiagramSpecToVisio.ps1`: core Visio generator.
- `scripts/Export-VisioPreview.ps1`: exports a generated `.vsdx` page to PNG.
- `scripts/Compare-DiagramPreview.ps1`: compares source and preview, producing JSON diff data and optional heatmap.
- `scripts/Build-VisioTemplateIndex.ps1`: indexes local `.vsdx` template libraries for reuse or style reference.

## Design Principles

- Prefer native editable Visio shapes over pasted full-image screenshots.
- Preserve legible text exactly when possible.
- Preserve connector direction, line style, and source-to-target topology.
- Use raster crops only for content that is hard to redraw cleanly, such as dense formulas, logos, 3D visual fragments, or detailed icons.
- Mark ambiguous edges as uncertain instead of silently guessing.
- Export and compare previews after material changes.

## License

MIT. See `LICENSE`.
