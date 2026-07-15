# Refined manifest schema

Use the refined manifest as the source of truth for final PPTX construction and delivery validation. Write identical logical records to CSV and JSON. Give every final PPTX object exactly one manifest row.

## Required fields

| Field | Allowed values or meaning |
|---|---|
| `slide_id` | Stable page identifier such as `slide01` or `image01` |
| `element_id` | Unique manifest identifier |
| `pptx_shape_name` | Unique final PowerPoint shape name; blank only for `backup_only` rows |
| `file_name` | PNG filename when an asset exists |
| `relative_path` | Path relative to the refined output root |
| `element_type` | Semantic category such as `title_txt`, `icon`, `figure`, `native_shape`, `background` |
| `final_representation` | `editable_text`, `native_shape`, `semantic_png`, `background`, or `backup_only` |
| `included_in_final` | `yes` or `no` |
| `decision_reason` | Specific reason for the chosen representation |
| `source_owner_id` | Unique owner for visible source content; required for included PNG assets |
| `x`, `y`, `width`, `height` | Source-pixel bounding box |
| `z_order` | Layer order from back to front |
| `opacity` | Estimated opacity, normally `1.0` |
| `rotation` | Rotation in degrees |
| `source_width`, `source_height` | Source-page dimensions |
| `pptx_left_emu`, `pptx_top_emu`, `pptx_width_emu`, `pptx_height_emu` | Final PPTX geometry |
| `text_readability` | `readable`, `unreadable`, or `not_text` |
| `recognized_text` | OCR or manually reconstructed text |
| `text_source` | `manual_transcription`, `trusted_ocr`, `user_provided`, or `not_applicable` |
| `text_accuracy_status` | `exact`, `approved`, `unreadable`, or `not_applicable` |
| `converted_to_editable_text` | `yes`, `no`, or `n/a` |
| `structure_complexity` | `simple`, `complex`, or `not_applicable` |
| `crop_mask_type` | `rectangle`, `circle`, `ellipse`, `irregular_mask`, or `full_background` |
| `crop_strategy` | Strategy used to derive the final crop |
| `alpha_strategy` | Strategy used to derive transparency |
| `transparent_background` | `yes`, `no`, or `n/a` |
| `alpha_note` | Explanation of alpha handling or exception |
| `crop_decision` | Final value must be `keep` or `not_applicable` |
| `review_status` | Final value must be `approved` or `not_applicable` |
| `review_note` | Specific closure note or exception approval |
| `needs_manual_review` | Final strict-delivery value must be `no` |
| `text_residue` | `yes`, `no`, or `not_applicable` |
| `simple_structure_residue` | `yes`, `no`, or `not_applicable` |
| `neighboring_content` | `yes`, `no`, or `not_applicable` |
| `clipped_content` | `yes`, `no`, or `not_applicable` |
| `duplicate_content` | `yes`, `no`, or `not_applicable` |
| `overlap_exception` | `yes`, `no`, or `not_applicable` |
| `conversion_note` | OCR, transcription, fallback, or background-cleaning notes |
| `source_page_path` | Path to the immutable source page used for identity comparison |
| `semantic_object_name` | Specific object or coherent composite represented by a semantic PNG; otherwise `not_applicable` |
| `semantic_scope` | `single_object`, `single_composite_visual`, or `not_applicable` |
| `semantic_page_tile_status` | `not_page_tile` or `not_applicable` |
| `semantic_independence_status` | `approved` or `not_applicable` |
| `semantic_audit_path` | Evidence image or audit sheet path for a semantic PNG; otherwise blank |
| `background_purity_status` | `pure_background`, `foreground_removed`, or `not_applicable` |
| `background_scene_complexity` | `complex`, `simple`, or `not_applicable` |
| `background_origin` | `image_gen_inpainted`, `image_gen_reconstructed`, `derived_background`, or `not_applicable` |
| `foreground_removal_method` | `image_gen`, `not_needed`, or `not_applicable` |
| `background_audit_path` | Source-cleaned-difference evidence path for a background; otherwise blank |
| `original_full_slide_reuse` | `yes`, `no`, or `not_applicable` |
| `icon_residue` | `yes`, `no`, or `not_applicable` |
| `shape_residue` | `yes`, `no`, or `not_applicable` |
| `chart_residue` | `yes`, `no`, or `not_applicable` |

## Representation invariants

- Require `final_representation=editable_text` for every row with `text_readability=readable` and `included_in_final=yes`.
- Require `final_representation=native_shape` for every row with `structure_complexity=simple` and `included_in_final=yes`.
- Require `text_residue=no`, `simple_structure_residue=no`, `neighboring_content=no`, `clipped_content=no`, and `duplicate_content=no` for every included `semantic_png`.
- Require `icon_residue=no`, `shape_residue=no`, and `chart_residue=no` for every included `semantic_png`.
- Require unique nonempty `source_owner_id` values for included `semantic_png` assets.
- Require every included `semantic_png` to name one object or coherent composite, set `semantic_scope` to `single_object` or `single_composite_visual`, set `semantic_page_tile_status=not_page_tile`, set `semantic_independence_status=approved`, and provide an existing `semantic_audit_path`.
- Require at least two included `semantic_png` rows on every slide that uses semantic PNG representation. Do not create fake assets to satisfy the count; use zero semantic PNGs when the source is fully reconstructable as editable text and native shapes.
- Reject rectangular semantic crops that collectively tile most of the page, even when each file has a unique name or owner ID.
- Require `transparent_background=yes` for included icons, logos, symbols, arrows, and irregular illustrations unless `alpha_note` contains an approved exception and `review_status=approved`.
- Require `included_in_final=no` and `final_representation=backup_only` for text PNG backups.
- Require a full-slide image covering at least 85% of the page to use `final_representation=background`.
- Prohibit an original, renamed, recompressed, or near-identical full-slide screenshot from the visible final deck regardless of representation.
- For an included `background` covering more than 95% of the page, require `background_purity_status=pure_background` or `foreground_removed`, `original_full_slide_reuse=no`, zero text/icon/shape/chart residue, an existing `background_audit_path`, and approved review evidence.
- Require `foreground_removal_method=image_gen` and an Image Gen origin for a complex background from which any foreground content was removed. Rebuild simple backgrounds as `native_shape` rather than a raster `background`.

## Crop lifecycle

Use `trim`, `expand`, `recrop`, or `manual_check` only during work in progress. Before strict delivery, resolve the condition, update `crop_decision` to `keep`, set `review_status=approved`, and preserve the history in `review_note`.

Do not declare a generic approval. Name what was checked, for example:

```text
review_note=expanded 8 px to restore arrow tip; checkerboard preview approved
```

## Large-background audit schema

When an included background covers more than 95% of a page, write `review/background_visual_audit/background_audit.csv` with one row per background and these fields:

- `element_id`
- `coverage_ratio`
- `background_purity_status`
- `background_origin`
- `foreground_removal_method`
- `original_full_slide_reuse`
- `text_residue`
- `icon_residue`
- `shape_residue`
- `chart_residue`
- `review_status`
- `evidence_file`

Set `evidence_file` to an existing source-cleaned-difference image. A manifest assertion without this audit does not prove that foreground content was removed.

## Recommended additional fields

Add these when available: `page_source_type`, `source_page_path`, `bbox_normalized`, `bbox_original`, `bbox_expanded`, `bbox_trimmed`, `font_family`, `font_size_pt`, `font_color`, `text_alignment`, `language`, `parent_group_id`, `shadow_effect`, and `overlap_with_ids`.
