---
name: visio-diagram-replica-workflow
description: Use when Codex needs to generate, replicate, QA, or prepare a Microsoft Visio diagram from a reference image, screenshot, whiteboard, system architecture diagram, flowchart, research workflow, business process, or when the user asks for Visio automation, editable .vsdx output, PNG/EMF preview export, visible Visio recording, or a reusable tutorial workflow. Covers reference-image decomposition, JSON drawing plans, direct Visio COM generation, sandbox cross-session proxy generation, stale-proxy fallback, preview-based visual verification, and recording preparation.
---

# Visio Diagram Replica Workflow

Use this skill to turn a reference diagram into an editable Visio `.vsdx`, with a preview image for fast visual QA. Default to automation first; use visible UI operations only when the user explicitly wants recording or manual demonstration.

## Operating Modes

Choose one mode and tell the user in one short sentence:

- **Automation first**: Build a JSON drawing plan, generate `.vsdx`, export a preview image, inspect the preview, then do targeted corrections.
- **Direct output**: If the user says "directly output" or "do not modify further", stop layout tuning and deliver the current `.vsdx` plus preview.
- **Visible recording**: Use Visio UI, mouse, and keyboard only after the user confirms recording has started.
- **Tutorial package**: Deliver the reusable plan, scripts, prompts, checks, and output structure so another user can repeat the workflow.

## Standard Workflow

1. Read the project `AGENTS.md` first. Keep final files in `outputs/`, temporary probes in `scratch/`, and reusable scripts in `scripts/`.
2. Check the reference image size. Use the image pixel width and height as the JSON page coordinate system unless another page size is required.
3. Decompose the diagram into zones, shapes, text labels, connectors, colors, line weights, dashed borders, and arrow directions.
4. Build a JSON plan in `outputs/<name>_plan.json`. Prefer stable coordinates over visual guesswork.
5. Generate Visio with the bundled script, resolving `scripts/...` relative to this `SKILL.md` unless the project already has a local copy:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/create_visio_from_plan.ps1 `
     -PlanPath outputs/<name>_plan.json `
     -OutVsdx outputs/<name>.vsdx `
     -OutPreview outputs/<name>.png
   ```

6. If the proxy is stale or the user wants the direct method, use direct COM:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/create_visio_direct_from_plan.ps1 `
     -PlanPath outputs/<name>_plan.json `
     -OutVsdx outputs/<name>.vsdx `
     -OutPreview outputs/<name>.png
   ```

7. Inspect the exported preview image. Fix only high-impact differences: missing shapes, wrong connector direction, major text overlap, or large zone misalignment.
8. Deliver the `.vsdx`, preview image, plan JSON, and any helper script used. Mention anything intentionally left approximate.

## Direct COM Fallback

Use direct COM when:

- `check_visio_proxy.ps1` says a lock exists but its PID is gone.
- `scratch/visio_proxy/commands/*.json` contains an old command and no response appears.
- The proxy process exists but does not consume commands.
- The user asks to stop refining and just output the drawing.

Stale proxy cleanup rule: remove only command files created by the current task and stale lock files whose PID no longer exists. Do not delete user source files or unrelated outputs.

The direct method is deterministic enough for most diagrams and avoids waiting on cross-session IPC. It should still export a preview image so the result can be checked quickly.

## JSON Plan Guidance

See `references/plan_schema_and_qa.md` for the supported schema and QA checklist. Core shape types are:

- `rect`: boxes and dashed zones.
- `oval`: ellipses.
- `line`: single straight connector.
- `polyline`: segmented connector; arrow is applied to the last segment.
- `text`: label without fill or border.

Use `RGB(r,g,b)` strings for colors. Use `dash: true` for dashed borders. Use `fontSize`, `bold`, `align`, and `textColor` for text. Prefer Arial for paper-style flowcharts.

## Preview QA Rules

- Export `.png` or `.emf` after every serious generation pass.
- Visually check page framing, dashed boundaries, connector endpoints, arrowheads, and label fit.
- If a label wraps badly, widen the text box first, then reduce font size slightly.
- If a user says "不要修改了" or "直接出图", stop tweaking and deliver the current export.
- Keep the editable `.vsdx` as the source of truth; the preview is only for QA and sharing.

## Built-In Scripts

| Script | Use |
|---|---|
| `scripts/check_visio_environment.ps1` | Check Visio, COM startup, export support, and proxy-related environment. |
| `scripts/create_visio_from_plan.ps1` | Main generator. Auto-routes to proxy in sandbox; supports `-PreferDirect` for direct COM when allowed. |
| `scripts/create_visio_direct_from_plan.ps1` | Thin direct-COM fallback wrapper around `create_visio_from_plan.ps1 -PreferDirect`. |
| `scripts/create_visio_via_proxy.ps1` | Execute a JSON plan through the cross-session proxy. |
| `scripts/visio_proxy_server.ps1` | Long-running interactive-user Visio COM proxy server. |
| `scripts/invoke_visio_proxy.ps1` | Send one proxy command and wait for the response. |
| `scripts/visio_proxy_client.py` | Python wrapper for proxy commands. |
| `scripts/check_visio_proxy.ps1` | Inspect proxy lock, status, Visio process, and queue health. |
| `scripts/start_visio_proxy.bat` | Start the proxy from the interactive Windows user. |
| `scripts/stop_visio_proxy.bat` | Stop the proxy. |
