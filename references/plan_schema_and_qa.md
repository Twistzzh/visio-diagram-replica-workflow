# Visio JSON Plan Schema And QA

Use this reference when building or debugging `outputs/<name>_plan.json` for `create_visio_from_plan.ps1`.

## Coordinate Model

- Use the reference image pixel size as `page.widthPx` and `page.heightPx`.
- Coordinates use image space: `(0,0)` is the top-left corner; `x` increases rightward and `y` increases downward.
- `scalePxPerInch` defaults to `100`; keep it stable across all shapes.
- Visio conversion happens inside the script, so JSON coordinates should stay in pixels.

## Minimal Plan

```json
{
  "page": {
    "name": "Research workflow replica",
    "widthPx": 1439,
    "heightPx": 725,
    "scalePxPerInch": 100
  },
  "referenceImage": "E:/temp/reference.png",
  "shapes": [
    {
      "type": "rect",
      "x1": 20,
      "y1": 80,
      "x2": 1410,
      "y2": 700,
      "text": "",
      "style": {
        "fill": "RGB(255,255,255)",
        "line": "RGB(80,80,80)",
        "weight": 1.2,
        "dash": true
      }
    }
  ]
}
```

## Supported Shape Types

- `rect`: rectangle or rounded-looking zone approximation.
- `oval`: ellipse or circular node.
- `line`: one straight connector, using `x1`, `y1`, `x2`, and `y2`.
- `polyline`: segmented connector, using `points`: `[[x,y], [x,y], ...]`.
- `text`: text-only box with no fill or border.

## Common Fields

- `type`: required shape type.
- `x1`, `y1`, `x2`, `y2`: bounding box for `rect`, `oval`, `text`, and `line`.
- `points`: ordered point list for `polyline`.
- `text`: optional label text.
- `arrow`: optional connector arrow: `begin`, `end`, or `both`; omit for no arrow.
- `style`: optional visual settings.

## Style Fields

- `fill`: Visio color formula, usually `RGB(r,g,b)`.
- `line`: Visio color formula for stroke.
- `weight`: line weight in points.
- `dash`: boolean dashed border or connector.
- `noFill`: boolean; hide fill.
- `noLine`: boolean; hide stroke.
- `fontSize`: text size in points.
- `bold`: boolean.
- `align`: paragraph alignment; `0` left, `1` center, `2` right.
- `textColor`: Visio color formula for text.

Prefer Arial for paper-style figures. The generator applies Arial automatically when styling text.

## Build And QA Checklist

1. Save the plan under `outputs/`.
2. Generate `.vsdx` and preview in one pass:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/create_visio_from_plan.ps1 `
     -PlanPath outputs/<name>_plan.json `
     -OutVsdx outputs/<name>.vsdx `
     -OutPreview outputs/<name>.png
   ```

3. Inspect the preview image before doing more edits.
4. Fix only high-impact issues first: missing regions, wrong arrow direction, severe overlap, or page framing.
5. For text wrapping problems, widen the box first; then lower `fontSize` by 0.5 to 1 pt.
6. For stale proxy or immediate-output requests, run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/create_visio_direct_from_plan.ps1 `
     -PlanPath outputs/<name>_plan.json `
     -OutVsdx outputs/<name>.vsdx `
     -OutPreview outputs/<name>.png
   ```

## Stale Proxy Fallback

Use direct COM when a proxy lock points to a missing PID, old commands remain unconsumed, or no response JSON appears after a reasonable wait. Clean up only files created by the current task or lock files whose PID no longer exists.
