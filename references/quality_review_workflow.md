# Quality review and delivery acceptance

Use this workflow for every refined slide. Complete the review against the PowerPoint-rendered preview, not only the construction canvas.

## 1. Check source and geometry

1. Confirm the source page in `source_pages/` is unchanged.
2. Confirm the PPTX slide ratio matches the source ratio within 0.1%.
3. Compare major region boundaries, object centers, spacing, corner radii, line weights, and z-order.
4. Reject unsupported stylistic additions such as new shadows, larger rounding, recoloring, or decorative elements.

## 2. Check text

1. Compare every readable title, body line, caption, label, legend, axis label, table entry, number, punctuation mark, and formula with the source.
2. Require an editable text box for every readable text region.
3. Inspect font substitution, CJK glyphs, font weight, alignment, line spacing, wrapping, clipping, and rotation in the rendered preview.
4. Set `text_accuracy_status=exact` when content matches the source. Use `approved` only for a documented intentional line-break or formatting adaptation.
5. Reject invented or uncertain wording.

## 3. Check native structure

1. Confirm panels, bands, borders, lines, dividers, arrows, badges, and simple symbols are native PowerPoint objects.
2. Reject simple structure embedded inside semantic PNG assets.
3. Compare fill, stroke, radius, shadow, and spacing with the source; do not accept unrequested redesign.

## 4. Check semantic PNG ownership

Open `review/crop_visual_audit/all_crop_audit.png` and inspect both source context and checkerboard preview for every final PNG.

Reject a crop when it:

- includes readable text or native-rebuild structure;
- includes a neighboring object, panel border, color band, or connector;
- clips the subject, arrow tip, chart endpoint, shadow, or anti-aliased edge;
- contains excessive background or transparent margin;
- has white/black halos, jagged edges, or deleted internal foreground;
- duplicates content owned by another crop;
- overlaps another semantic ownership box without an approved exception.

Resolve each item to `crop_decision=keep` and `review_status=approved`. Strict delivery requires zero unresolved items.

Also reject a semantic PNG when it is a page strip, column, quadrant, panel screenshot, or other rectangular page tile rather than one named semantic object. When semantic PNGs are used on a slide, require multiple independent assets and verify that their combined rectangles do not form a near-complete page mosaic.

## 5. Check large backgrounds

For every included background covering more than 95% of the page:

1. Open the source-cleaned-difference evidence in `review/background_visual_audit/`.
2. Confirm the visible asset is not the original source page, a renamed copy, a recompressed copy, or a near-identical full-page derivative.
3. Confirm the background is pure or has no remaining text, icons, shapes, charts, or other foreground content.
4. Confirm a complex photographic, textured, or illustrative background was cleaned or reconstructed with Image Gen.
5. Reject pseudo-text, ghosts, generated false details, damaged subjects, duplicated foreground, and unsupported layout changes.
6. Confirm simple backgrounds were rebuilt with native PowerPoint shapes instead of raster cleaning.

## 6. Check the actual PowerPoint render

Export every final slide through PowerPoint or an equivalent faithful renderer. Check:

- aspect ratio and output dimensions;
- out-of-bounds shapes;
- clipped text and images;
- accidental overlaps and duplicated text;
- duplicate PNG content;
- layout drift and altered whitespace;
- transparency against the real slide background;
- CJK and formula rendering;
- shadows, soft fades, and z-order.

Use the diff image to locate changes, not as the sole pass/fail signal. Rasterized editable text can produce large pixel differences despite correct reconstruction.

## 7. Write the per-slide report

Write these required fields to `quality_report/quality_report.csv` and JSON:

| Field | Pass condition |
|---|---|
| `slide_id` | Stable page identifier |
| `slide_pass` | `yes` only when every condition below passes |
| `preview_status` | `approved` |
| `text_accuracy_status` | `exact` or documented `approved` |
| `aspect_ratio_status` | `match` |
| `unresolved_review_count` | `0` |
| `crop_failure_count` | `0` |
| `duplicate_asset_count` | `0` |
| `overflow_count` | `0` |
| `clipping_count` | `0` |
| `overlap_count` | `0` |
| `background_gate_status` | `approved` or `not_applicable` |
| `original_full_slide_status` | `not_reused` |
| `semantic_split_status` | `approved` or `not_applicable` |
| `semantic_tile_status` | `not_page_tiles` or `not_applicable` |
| `quality_notes` | Specific review summary and approved exceptions |

Do not replace these fields with a prose-only assertion.

## 8. Run strict validation

```text
python scripts/validate_delivery.py <refined_output_dir> --pptx <refined_deck.pptx>
```

Inspect `quality_report/delivery_validation.csv` when validation fails. Fix the deck or manifest and rerun. Deliver only after the script exits with code 0.
