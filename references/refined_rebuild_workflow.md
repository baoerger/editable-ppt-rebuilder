# Refined editable rebuild workflow v2.1

Use this workflow as the default final-output path for English, Chinese, mixed-language, visual, academic, technical, poster, and screenshot-based raster slides.

## Contents

1. Baseline discovery
2. Output setup
3. Region map and ownership plan
4. Representation decisions
5. Text reconstruction
6. Native structure reconstruction
7. Semantic PNG extraction
8. Full-background foreground removal
9. Deterministic build pattern
10. Review closure
11. Delivery validation
12. Batch work

## 1. Baseline discovery

Run baseline decomposition unless the user explicitly requests planning or extraction only:

```text
python scripts/decompose_visual_elements.py input.png \
  --outdir baseline_output \
  --dpi 300 \
  --granularity fine \
  --ocr \
  --ocr-lang chi_sim+eng \
  --ocr-confidence-threshold 75 \
  --editable-text \
  --review \
  --quality-check
```

Use baseline output to discover text regions, visual candidates, page dimensions, and likely structure. Do not treat baseline OCR or its recomposed deck as final.

Check OCR language support. Use unsupported-language OCR only for rough region discovery.

## 2. Output setup

Create the refined output bundle before building. Preserve the source page unchanged and keep baseline output separate.

Use a task-specific deterministic rebuild script when it improves reproducibility. Write all refined assets directly under the refined output root.

Set the PowerPoint slide dimensions from the source ratio. Do not force 16:9, 4:3, or another standard ratio when the source differs.

## 3. Region map and ownership plan

Map these regions in source-pixel coordinates:

- page background and decorative framing;
- title, subtitle, section headers, body text, captions, and footers;
- panels, cards, bands, table frames, and chart frames;
- axes, bars, ticks, labels, legends, and connectors;
- people, products, photos, logos, icons, diagrams, plots, and scientific figures;
- badges, callouts, symbols, and decorative linework.

Assign one `source_owner_id` to each visible source object. Do not let two final PNG assets own the same semantic object or source pixels. Record intentional overlaps explicitly.

Do not plan the page as a mosaic of rectangular screenshot regions. A semantic ownership box must name one object or one coherent composite visual, not a top strip, column, quadrant, card-sized screenshot, or other page tile.

## 4. Representation decisions

Choose exactly one final representation for every object:

- `editable_text`: all readable text and simple legible formulas;
- `native_shape`: simple structure and geometric symbols;
- `semantic_png`: complex or style-sensitive visual cores;
- `background`: pure or approved text-removed background;
- `backup_only`: hidden recovery assets such as text PNG backups.

Record the decision and a specific reason in the refined manifest before final QA.

Examples:

| Source object | Final representation | Reason |
|---|---|---|
| Chinese title | `editable_text` | Readable text has high editing value |
| Rounded panel | `native_shape` | Simple structure can be reproduced faithfully |
| Network diagram | `semantic_png` | Exact topology and spatial relationships matter |
| Photo with overlaid title | `background` plus `editable_text` | Preserve scene while making title editable |
| Unreadable formula | `backup_only` or reviewed PNG | Transcription would invent content |

## 5. Text reconstruction

Use the source image or user-provided text as ground truth. Use OCR only when visibly correct.

For every readable text region:

1. Transcribe exact wording, numbers, punctuation, and notation.
2. Preserve line order and source line breaks unless rendered PowerPoint requires a documented adaptation.
3. Create an editable PowerPoint text box at the source position.
4. Match font family, weight, size, color, alignment, spacing, and rotation as closely as practical.
5. Export the final PPTX and inspect actual rendering.
6. Set `text_accuracy_status=exact` or document an approved adaptation.

For Chinese text, prefer the deck font when known; otherwise use a common CJK font such as `Microsoft YaHei` or `SimHei`. Use conservative sizes and wider boxes because PowerPoint may render CJK glyphs larger than expected.

Do not invent unreadable content. Keep it as a reviewed PNG fallback and fail strict delivery until approved.

## 6. Native structure reconstruction

Rebuild page backgrounds, panels, rectangles, rounded rectangles, pills, color bands, straight lines, dividers, connectors, arrows, borders, badges, simple chart axes, simple bars, and geometric symbols as native PowerPoint objects.

Use source colors and geometry. Do not add shadows, larger corner radii, heavier outlines, or spacing changes merely to make the deck look cleaner.

Name each PowerPoint object deterministically and write the same name to `pptx_shape_name` in the manifest.

## 7. Semantic PNG extraction

Create semantic candidates only after subtracting readable-text masks and native-structure masks. When semantic PNG representation is used on a slide, produce multiple independent semantic PNGs; a single large residual screenshot is not a valid split.

For each candidate:

1. Expand the seed box slightly to protect edges, shadows, and arrow tips.
2. Estimate the local background from pixels connected to the crop border.
3. Remove only external background; preserve internal white or light foreground.
4. Clean the alpha mask lightly and preserve anti-aliasing.
5. Trim to visible bounds and add small transparent padding.
6. Check for text residue, simple-structure residue, neighboring content, clipping, duplicate ownership, and unexplained overlap.
7. Generate a source-context and checkerboard audit preview.
8. Resolve the crop to `crop_decision=keep` and `review_status=approved`.

Record `semantic_object_name`, `semantic_scope`, `semantic_page_tile_status`, `semantic_independence_status`, and `semantic_audit_path`. Set `semantic_scope` only to `single_object` or `single_composite_visual`. Reject any group of rectangular crops whose union reconstructs most of the page as a tile mosaic.

Keep rectangular crops for photos, dense plots, screenshots, or cases where transparency would damage fidelity. Record the fallback and approve it explicitly.

Do not crop a full card merely because it contains a complex icon. Crop the visual core, then rebuild the card and its text separately.

## 8. Full-background foreground removal

Use background cleaning only for slides dominated by a photographic, textured, or illustrative background with embedded foreground content.

Preserve the original source. Never place that source page, a recompressed copy, or a near-identical full-page derivative in the visible final deck.

Classify the background before cleaning:

- Rebuild solid colors, ordinary gradients, simple lines, and geometric background structure with native PowerPoint shapes.
- Use Image Gen semantic inpainting or reconstruction for complex photographic, textured, shaded, or illustrative backgrounds. Remove all text, icons, shapes, charts, and other foreground objects while preserving the background scene, lighting, materials, textures, composition, and color grading.

Reconstruct readable text as editable text, simple foreground structure as native shapes, and complex foreground visuals as independent semantic PNGs.

If an included background covers more than 95% of the page, produce `review/background_visual_audit/background_audit.csv` and a source-cleaned-difference evidence image. Explicitly verify `text_residue=no`, `icon_residue=no`, `shape_residue=no`, `chart_residue=no`, `original_full_slide_reuse=no`, and `review_status=approved`.

Inspect the final render for ghosts, pseudo-characters, generated false details, missing objects, damaged faces or products, duplicated content, and layout drift. Regenerate damaged background areas or fail the slide. Do not repair the background by pasting the original full-page screenshot back into the visible deck.

Record the background as `final_representation=background`, identify its origin and Image Gen method, and set `review_status=approved` only after visual inspection. Image Gen completion alone does not satisfy the gate.

## 9. Deterministic build pattern

Use these script components:

1. source and output path constants;
2. source width, source height, slide width, and slide height;
3. pixel-to-EMU conversion helpers;
4. helpers for text, shapes, lines, arrows, and pictures;
5. unique PowerPoint shape naming;
6. semantic crop and alpha functions;
7. refined manifest CSV/JSON writing;
8. crop audit CSV/JSON and visual sheets;
9. semantic-independence and background-cleaning audit evidence;
10. review overlay generation;
11. PowerPoint preview export;
12. per-slide quality report writing;
13. strict delivery validation.

Give every final PowerPoint object one manifest row. Keep backup assets in the manifest with `included_in_final=no`.

## 10. Review closure

Perform three passes:

### Pass A: structural inspection

Open the PPTX with `python-pptx`. Confirm slide count, aspect ratio, shape names, object counts, and zero out-of-bounds objects.

### Pass B: visual inspection

Compare source, PowerPoint render, diff, refined overlay, and crop audit. Fix text, geometry, crop, transparency, duplication, overlap, and z-order defects.

### Pass C: record closure

Update manifest and crop audit rows. Require zero `manual_check`, `recrop`, `trim`, `expand`, or `needs_manual_review=yes` rows before strict delivery.

## 11. Delivery validation

Write the required `quality_report.csv` fields described in `quality_review_workflow.md`, then run:

```text
python scripts/validate_delivery.py refined_output \
  --pptx refined_editable.pptx
```

Fix every reported error. Do not deliver on a warning-style prose judgment when the validator reports failure.

## 12. Batch work

Use `scripts/run_batches.py` for deterministic baseline processing and preserve original input order. Curated rebuilds may remain task-specific.

Validate each refined item before merging. Exclude failed items from a final merge and report them clearly. A successful baseline process does not satisfy refined delivery acceptance.
