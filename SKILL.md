---
name: codeximage-ppt
description: >-
  Rebuild image-based PPT/PPTX files and slide screenshots into source-faithful editable PowerPoint decks with mandatory delivery validation. Use when converting flattened raster slides into editable text, native PowerPoint structure, and ImageGen-regenerated high-resolution visual cores; when every small icon must follow a reviewed ImageGen green-screen, local alpha-removal, transparency-check, embedding, and audit workflow; when OCR may be incomplete or unreliable; when crop ownership, transparency, visual fidelity, or duplicate-content risks require review; or when a final deck must fail closed unless text editability, native structure, crop integrity, review closure, and PowerPoint-rendered QA all pass. Use scripts/run_batches.py for repeated baseline processing and scripts/validate_delivery.py before final delivery.
---

# CodexImage-PPT

## Purpose

Rebuild raster slides into usable editable PowerPoint decks. Treat the source page as visual and textual ground truth. Do not claim recovery of original vector objects.

Produce a refined editable rebuild as the primary result. Treat the automated OCR recomposition as a reference and backup only.

Preserve the original source page unchanged. Reconstruct each visible object using exactly one final representation:

- `editable_text` for readable text;
- `native_shape` for simple presentation structure;
- `semantic_png` for complex or style-sensitive visual content;
- `background` for a pure or approved text-removed background;
- `backup_only` for non-visible recovery assets.

## Required references

Read the following files before the corresponding work:

- Read `references/refined_rebuild_workflow.md` before building a refined deck.
- Read `references/manifest_schema.md` before writing the refined manifest.
- Read `references/quality_review_workflow.md` before final QA and delivery.
- Read `references/imagegen_visual_core_workflow.md` before using ImageGen for a background, icon, product, vehicle, or illustration.

## Core workflow

1. Extract an embedded full-slide raster directly when possible. Otherwise render at 240-300 DPI. Write immutable source pages and an image source report.
2. Run `scripts/decompose_visual_elements.py` as a baseline discovery step with OCR, review overlays, and quality output enabled.
3. Inspect the source, baseline manifest, overlay, and rough preview. Do not trust OCR that is visibly wrong or unsupported for the slide language.
4. Map every visible region and classify every object with the mixed-fidelity gate.
5. Build the refined deck deterministically from source coordinates. Preserve the source aspect ratio.
6. Export readable text as editable PowerPoint text boxes and keep text PNGs only as `backup_only` assets.
7. Rebuild simple structure as native PowerPoint objects.
8. Regenerate the visual core of every small icon with ImageGen using the mandatory green-screen visual-core workflow. Export the resulting independent transparent PNG only after local alpha removal and checkerboard review. Keep readable text and simple neighboring structure out of each icon. Never substitute rectangular page tiles for semantic objects.
9. When a complex full-page background must be cleaned or reconstructed, use Image Gen semantic inpainting or reconstruction and preserve review evidence. Rebuild simple backgrounds with native PowerPoint shapes.
10. Generate crop, semantic-independence, and full-background audit evidence and close every review item.
11. Export the refined PPTX through PowerPoint or an equivalent faithful renderer and inspect the actual rendered result.
12. Write the required manifest and per-slide quality report.
13. Run `scripts/validate_delivery.py`. Deliver the refined deck only when the validator exits with code 0.

Skip the refined path only when the user explicitly requests automatic decomposition, planning, or raw element extraction. Label such output as non-final.

## Mixed-fidelity decision gate

Classify each visible object by editing value and redraw risk.

Use `editable_text` for all readable titles, subtitles, body text, labels, captions, legends, axis labels, table text, footers, callouts, and simple formulas.

Use `native_shape` for backgrounds, panels, rectangles, rounded rectangles, pills, badges, color bands, straight lines, dividers, connectors, arrows, table/grid borders, simple status bars, basic chart scaffolds, and clean geometric symbols.

Use `semantic_png` for photos, people, products, logos, shaded illustrations, scientific schematics, complex icons, dense plots, clusters, networks, topology diagrams, model architectures, composite mini-diagrams, and any visual whose meaning or style could be damaged by redrawing.

### Mandatory ImageGen small-icon workflow

Treat every small visual icon as an ImageGen regeneration task. This is a hard instruction, not a quality-dependent suggestion: do not crop, upscale, trace, or reuse a small source icon as a visible final icon. Apply it to all pictograms, equipment symbols, vehicles, people, product miniatures, badges, generic logos, diagram glyphs, decorative illustrations, and any other compact visual mark.

For every small icon, perform this complete sequence:

1. Separate the icon's non-text visual core from surrounding panels, arrows, labels, borders, badges, and data.
2. Generate one high-resolution visual core with ImageGen on a perfectly flat chroma-key background. Do not ask ImageGen to draw slide text, numbers, arrows, label wording, QR content, or legend content.
3. Remove the chroma key locally with the installed ImageGen helper, preserving an RGBA asset with transparent corners.
4. Inspect the asset on a checkerboard for fringe, halo, clipped subject, key-color residue, invented text, and unintended extra objects. Regenerate or reprocess until approved.
5. Embed the approved transparent PNG in the PPTX and reconstruct all surrounding text and simple geometry as native editable PowerPoint objects.
6. Copy the final asset to `split_png_elements/`, generate its crop audit, and record a manifest/audit row for every placed icon instance.

If an icon contains readable text, an official mark, or a precise logo, still regenerate its graphic visual core as required, but remove the readable or exact mark from the generated prompt and rebuild it separately as editable text or a native shape. Do not accept generated pseudo-text or a generated approximation as the authoritative exact mark. When no editable reconstruction is possible, flag the mark for manual review rather than leaving source pixels inside the icon.

For backgrounds, continue to use ImageGen whenever a complex scene must be reconstructed or cleaned; use native shapes for simple backgrounds. For all eligible objects use this order: generated text-free visual core → editable text → native labels, pills, borders, connectors, arrows, rings, legend lines, and cards.

Generate each icon type with a focused prompt specifying subject, viewpoint, palette, intended slide use, and exclusions. Generate transparent-ready assets on a flat chroma-key background, remove the key with the installed helper, inspect an alpha checkerboard preview, and only then embed the PNG. When an icon repeats at several locations, reuse the approved generated type in PowerPoint but record one independent manifest/audit ownership row per placed instance; use stable instance asset copies when the delivery validator requires unique asset ownership.

Record `background_origin=image_gen_reconstructed` or an ImageGen-specific `conversion_note`, `foreground_removal_method=image_gen` when applicable, prompt summary, source/semantic role, alpha method, and review evidence. Do not claim the generated asset is an exact copy of the source; describe it as a reviewed visual-core reconstruction.

Do not redraw a complex semantic visual merely because it is technically possible. When uncertain, use editable text, native containers, and a semantic PNG visual core.

Record `final_representation`, `decision_reason`, `included_in_final`, and `review_status` for every manifest row.

## Mandatory delivery gates

Treat every gate in this section as blocking. Do not soften `must` into a recommendation and do not mark a failed slide complete.

### Gate 1: source integrity and faithful reconstruction

- Keep every source page unchanged in `source_pages/`.
- Match the source and PPTX aspect ratios within 0.1%.
- Preserve readable wording, numbers, punctuation, mathematical notation, and ordering exactly.
- Preserve source line breaks unless PowerPoint rendering requires a change; record every intentional change.
- Preserve geometry, hierarchy, colors, corner radii, line weights, shadows, and spacing as closely as practical.
- Do not add shadows, decoration, rounded corners, icons, colors, or other styling unsupported by the source.

### Gate 2: representation compliance

- Convert 100% of readable text to editable text boxes.
- Rebuild 100% of simple structure as native PowerPoint objects.
- Regenerate 100% of small visual icons with ImageGen visual cores using the mandatory green-screen → local alpha removal → checkerboard transparency review → PPT embedding → audit workflow. Do not visibly reuse a source small-icon crop, including when it is sharp enough to reuse.
- Keep complex semantic visuals as independent PNG assets when redraw risk is material.
- Do not leave readable text inside a final semantic PNG.
- Never include the original full-slide screenshot in the visible final reconstruction, including when renamed, recompressed, placed behind editable text, or declared as `background`.
- When an included `background` covers more than 95% of the page, require proof that it is a pure background or that all text, icons, shapes, charts, and other foreground content were removed. Fail when proof is absent or residue remains.
- Use Image Gen semantic inpainting or reconstruction for foreground removal on complex photographic, textured, or illustrative backgrounds, following the installed `imagegen` skill instructions when this gate triggers. Rebuild simple solid, gradient, line, or geometric backgrounds with native PowerPoint shapes instead of raster cleaning.
- Keep unreadable text as PNG only with `text_readability=unreadable`, `needs_manual_review=yes`, and a user-visible delivery warning. Such a slide fails strict delivery until approved.

### Gate 3: crop integrity and ownership

- Give every visible source object one owner in the final deck.
- Do not include the same source pixels or semantic object in multiple final PNG assets.
- Do not let semantic PNGs contain neighboring labels, panel borders, color bands, connectors, or unrelated objects.
- Require multiple included semantic PNGs on a slide whenever semantic PNG representation is used. A slide with exactly one semantic PNG fails strict delivery.
- Require every semantic PNG to represent one named object or one coherent composite visual. Reject page strips, panel-sized screenshot crops, quadrant crops, and any collection of rectangular crops that reconstructs the page by tiling.
- Require `split_png_elements/` to contain the independent assets referenced by the manifest; files that merely rename or re-encode page tiles do not count.
- Do not accept clipped subjects, excessive margins, page-background residue, white or black halos, jagged edges, or deleted internal foreground.
- Require transparent external backgrounds for icons, logos, symbols, arrows, and irregular illustrations unless an approved alpha exception explains why transparency would damage fidelity.
- Require `overlap_exception=yes` plus an approved review note for intentional overlap between final semantic PNG ownership boxes.

### Gate 4: review closure

- Generate `review/crop_visual_audit/crop_audit.csv` and checkerboard previews for every final non-background PNG.
- Resolve `trim`, `expand`, `recrop`, and `manual_check` before delivery.
- Set final `crop_decision=keep` or `not_applicable` and `review_status=approved`.
- Require zero rows with `needs_manual_review=yes` in strict delivery.
- Preserve reviewer notes for exceptions instead of replacing them with generic statements.

### Gate 5: rendered PowerPoint QA

- Export a preview from the final PPTX, not from the construction canvas.
- Require zero out-of-bounds objects, clipped text, clipped images, accidental overlaps, duplicated text, duplicated assets, and unexplained layout drift.
- Inspect font substitution, CJK rendering, wrapping, line spacing, formulas, z-order, transparency, and shadows.
- Reject a preview whose dimensions or aspect ratio do not match the source page.

### Gate 6: auditable pass/fail report

Write one row per slide in `quality_report/quality_report.csv` with at least:

- `slide_id`
- `slide_pass`
- `preview_status`
- `text_accuracy_status`
- `aspect_ratio_status`
- `unresolved_review_count`
- `crop_failure_count`
- `duplicate_asset_count`
- `overflow_count`
- `clipping_count`
- `overlap_count`
- `quality_notes`

Set `slide_pass=yes` only when all counters are zero, text accuracy is `exact` or `approved`, aspect ratio is `match`, the final render is approved, and all review items are closed.

Run:

```text
python scripts/validate_delivery.py <refined_output_dir> --pptx <refined_deck.pptx>
```

Treat a nonzero exit code as a failed delivery. Do not write “no severe issues” when unresolved rows or failed counters remain.

## OCR capability gate

Check available OCR languages before trusting recognized text. If Chinese text is present without `chi_sim`, `chi_tra`, or another suitable model, use OCR only for region discovery and reconstruct readable text manually from the source.

Reject OCR text when it is gibberish, merged across unrelated regions, assigned to the wrong object, clipped, badly wrapped, or visibly inconsistent with the source. Record `ocr_language_missing_*`, `ocr_gibberish_replaced`, or `manual_text_reconstruction` as appropriate.

Never invent unreadable wording.

## Full-background foreground-removal gate

Apply this gate to poster, hero, title, product, venue, portfolio, or editorial slides when at least two conditions hold:

- a photo or illustration covers at least 85% of the slide;
- baseline decomposition merges text with the background;
- decorative raster text overlays complex imagery;
- OCR is missing or unreliable;
- simple local inpainting would leave obvious ghosts.

When triggered:

1. Preserve the original source page.
2. Identify every text region that must become editable.
3. Classify the background as simple or complex. Rebuild a simple background with native PowerPoint shapes. For a complex background, use Image Gen semantic inpainting or reconstruction to remove all text, icons, shapes, charts, and other foreground content while preserving the non-foreground scene, product, lighting, texture, composition, and color grading.
4. Reconstruct all readable text as editable text boxes.
5. Reconstruct icons and complex figures as independent semantic PNGs, and simple structures as native PowerPoint objects.
6. Produce a background audit showing the immutable source, cleaned background, difference view, method, origin, coverage, and explicit residue checks for text, icons, shapes, and charts.
7. Inspect for ghosts, pseudo-characters, duplicated content, invented generated details, damaged subjects, and layout drift.
8. Restore damaged non-text regions through a reviewed regeneration or fail the slide for manual review. Do not paste the original full-page screenshot back into the visible deck.

Do not use an unreviewed generated background in the final deck. Image Gen use is necessary for complex semantic cleaning but is not proof by itself; the audit and rendered PowerPoint review must also pass.

## Crop and alpha rules

Start from a visual candidate after subtracting text and native-structure masks. Expand slightly to protect anti-aliased edges, estimate the local border-connected background, remove only external background, clean the alpha mask lightly, trim to visible bounds, and add small transparent padding.

Prefer these strategies:

- `border_connected_background_flood` for icons, logos, people, illustrations, and charts on plain backgrounds;
- `near_background_alpha` only when it preserves internal light foreground;
- `segmentation_mask_alpha` for manually or visually refined irregular objects;
- `rectangular_crop_fallback` for photos, dense plots, screenshots, or approved cases where alpha would damage fidelity.

Record `crop_strategy`, `alpha_strategy`, `alpha_note`, `crop_decision`, `review_status`, and all residue/duplication checks defined in the manifest schema.

## Refined output contract

Produce this bundle for final work:

```text
<name>_refined_editable.pptx
<name>_refined_editable_output/
|-- source_pages/
|-- split_png_elements/
|-- split_png_elements.zip
|-- visual_elements_manifest.csv
|-- visual_elements_manifest.json
|-- image_source_report.csv
|-- image_source_report.json
|-- review/
|   |-- <page>_refined_elements_overlay.png
|   |-- background_visual_audit/
|   |   |-- background_audit.csv
|   |   `-- <page>_<background>_audit.png
|   `-- crop_visual_audit/
|       |-- crop_audit.csv
|       |-- crop_audit.json
|       `-- all_crop_audit.png
|-- quality_preview/
|   `-- powerpoint_export/
|-- quality_report/
|   |-- quality_report.csv
|   |-- quality_report.json
|   |-- <page>_original.png
|   |-- <page>_recomposed.png
|   |-- <page>_diff.png
|   |-- delivery_validation.csv
|   `-- delivery_validation.json
`-- recomposed_from_elements.pptx
```

Keep the baseline output separately and never present it as the primary refined result.

## Implementation pattern

Use a deterministic rebuild script with source dimensions, pixel-to-EMU helpers, reusable text/shape/picture helpers, explicit z-order, manifest writing, crop and background audit generation, PowerPoint preview export, and validation.

Use source coordinates throughout. Give each final PPTX object exactly one manifest row and assign a stable `pptx_shape_name` so the validator can compare the manifest with the deck. Record semantic object identity and independence evidence for every included semantic PNG. Record background origin, Image Gen cleaning method, residue status, and audit evidence for every included large background.

For Chinese text, prefer a deck-consistent CJK font such as `Microsoft YaHei` or `SimHei`. Use conservative font sizes and explicit line breaks only when confirmed in the rendered preview.

## Batch processing

Use `scripts/run_batches.py` for repeated baseline decomposition. Keep batch instructions, logs, summaries, and merge order deterministic.

Do not treat a successful baseline batch as a successful refined delivery. Validate each curated refined output separately. Merge only decks whose delivery validator passed, and preserve original input order.

## Limitations

Raster input cannot perfectly recover original vectors, hidden text, chart data, or object semantics. Prefer a faithful independent PNG over an inaccurate native redraw. Record uncertainty explicitly and fail closed when a mandatory gate cannot be satisfied.
