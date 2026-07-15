# ImageGen visual-core regeneration

Use this mandatory workflow for every small visual icon. Do not crop, upscale, trace, or visibly reuse a small source icon in the final deck.

## 1. Decide what may be regenerated

Regenerate every small visual icon: pictogram, equipment symbol, vehicle, person, product miniature, badge, generic logo, diagram glyph, or decorative illustration. Treat the source only as visual guidance; do not visibly reuse its small-icon pixels.

When a small icon includes readable text, an official mark, or a precise logo, regenerate only the non-text graphic core and rebuild the readable/exact content separately with editable text or native shapes. Never treat generated pseudo-text or a generated approximate logo as exact evidence. If exact content cannot be reconstructed, mark it for manual review.

## 2. Separate visual core from editable content

Before prompting, identify the visual core and its surrounding editable objects. Keep all readable text, labels, numbers, arrows, rings, routes, legend entries, callout panels, borders, and simple geometry out of the generated asset.

For a background, request a clean no-text scene with deliberate negative space. For every icon, request exactly one centered subject. Do not generate a whole card or a page fragment.

## 3. Generate and make transparency

Follow the installed `imagegen` skill. Use its built-in tool by default. For every icon, request:

```text
Scene/backdrop: perfectly uniform #00ff00 chroma-key background, no floor or gradient.
Constraints: no text, no labels, no watermark, no shadow, no reflection, no extra objects.
```

Then use the installed helper:

```text
python "$CODEX_HOME/skills/.system/imagegen/scripts/remove_chroma_key.py" \
  --input <generated-chroma.png> --out <asset.png> \
  --auto-key border --soft-matte --transparent-threshold 12 \
  --opaque-threshold 220 --despill
```

Validate that the result is RGBA, has transparent corners, has no visible key-color fringe, no invented text or mark, and includes the entire subject. Retry once with `--edge-contract 1` only if a thin fringe remains; regenerate if content or silhouette is wrong.

## 4. Insert and audit

Save project-bound assets under the task workspace, then copy final assets to `split_png_elements/`. Keep images embedded in the PPTX. Generate checkerboard asset previews and record one crop-audit row per placed semantic PNG with:

- `element_id`, `crop_decision=keep`, `review_status=approved`;
- `transparent_background=yes`, alpha strategy, and review note;
- `text_residue=no`, `simple_structure_residue=no`, `neighboring_content=no`, `clipped_content=no`, and `duplicate_content=no`;
- a specific semantic object name and audit path.

For a repeated icon, retain the approved visual design but record one audit/ownership row per location. Use a stable instance copy in `split_png_elements/` whenever strict delivery validation needs distinct asset ownership.

For a generated full-page background, perform the normal large-background audit and explicitly state that the original full-slide raster is not reused.

## 5. Final check

Check the PowerPoint-rendered preview for matte halos, scale mismatch, accidental clipping, weak contrast, and duplicate or stale source icons. Keep all surrounding business content editable. Run the standard delivery validator; resolve manifest, crop-audit, and asset-ownership errors before delivery.
