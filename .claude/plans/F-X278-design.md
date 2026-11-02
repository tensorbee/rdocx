# F-X278, Complete Word rendering fidelity and font policy

**Status**: approved

**Sprint**: S96

**Size**: L

**Depends on**: none

**Test gate**: differential

Root adopts the accepted V6 preparation as this canonical draft on 2026-11-01. The preparation is `/private/tmp/rdocx-grok-harness/f-x278-preparation/GROK_PROPOSAL_V6.md` at SHA-256 `c1b425a1076d2a7d1e6c74bdd27453498c0fdcf6fa9e7d0821e56bf62fae6f94`, with `HANDOFF_V6.md` at SHA-256 `0bf0b96d2729bf2469c3b4dc0f60480af9bbdc6804693e71a8200de3cee7953c`. Root final 3 is `.claude/reviews/F-X278-codex-preparation-final-3.md` at SHA-256 `875c0ac092d6d2e92b008acc702c57d80433615a68e27055a79532044ac68069`. Prior final 2 substantive dispositions remain applicable. Q-S96-07 is answered as bundled default, system fallback opt-in, and caller fonts highest. Three design peers are used and zero remain. A later implementation has its own cap of three. This adoption is not approval, a claim, implementation, or a counter reset. Source F-X272 remains rejected on C1 and C4 and paused for design clarification. It is not a provider. F-X276 remains draft. Registration already lists this story as pending and unclaimed, so this adoption does not edit trackers.

## Operative rolling approval, 2026-11-01

Root approves this contract for the user's explicit rolling parallel start after the recorded final design reviews and canonical fidelity checks. Earlier preparation and batch-blocking wording is historical. The immutable product source Base is reviewed commit `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, excluding the preserved unfinished F-X272 working changes. No product dependency exists. Implementation must recheck exact source locations on its worker Base and preserve every acceptance criterion, signature, HLD impact and pending final rider. Q-S96-07 is answered as bundled fonts by default, system fallback opt-in and caller fonts highest. This is contract approval, not implementation acceptance or a provider checkpoint. Three design peers remain used. Implementation has a separate cap of three independent passes.


## Problem

Issue 433 is seven rendering outcomes plus one font-policy split. The font split is entry-point selection. It is not a broken caller-priority rule inside one engine. Citations below are frozen 919-packet pins unless Evidence pins says a line was re-read in the current worktree.

1. TOC ink. `insert_toc` writes paragraph style id `TOC{level}` and name `TOC {level}` at `crates/rdocx/src/field.rs:10395-10420`. Run style is resolved at `crates/rdocx-layout/src/engine.rs:11822-11829`. Stored segment style is resolved at `:12331-12336`. A TOC entry whose run uses the Hyperlink character style still takes that colour and underline.
2. First-line indent. `line.rs:2256-2260` and `:2307-2310` prefer `ind_hanging > 0` over `ind_first_line`. The four issue cases fail when an earlier layer leaves the other field set. `CT_PPr::merge_from` and `merge_direct_ppr` do not clear the absent sibling.
3. Table borders. `resolve_table_style_cell` copies style `tblBorders` onto cell borders at `table.rs:2165-2172`. Paint then prefers the cell edge. `resolved_cell_edge` at `table.rs:983-992` treats a cell `none` as unset, so copying `none` does not clear a promoted grid.
4. Row split. `split_simple_table_row` returns the whole row when `cell.v_align.is_some()` at `paginator.rs:1844`. That line was re-read in the current worktree and still contains the bail.
5. Automatic colour. `resolve_run_color` returns `Color::BLACK` at `engine.rs:16529`. That return was re-read and is still the `auto` fallback. Cell fill is known at `table.rs:690-697` and is not passed into `layout_cell_content` at `:755`.
6. Wingdings 2. Shared `resolve_symbol_text` at `font.rs:1911` maps Wingdings and Symbol only. `rpptx-render` calls that shared function.
7. Font policy. With no caller font, Python `page_count` and `layout` call `layout_deterministic` (`rdocx-py` `document.rs:12510-12512` and `:15783`). Python `to_pdf` calls normal `to_pdf`. CLI `convert` loads fonts at `commands.rs:811-817` and lays out PDF with bundled fallback at `:833`. CLI `layout` uses bundled fallback at `:438`. CLI `render` uses bundled fallback at `:2248-2253`. A non-empty caller font makes Python `page_count` and `layout` call `layout_with_additional_fonts_and_options`. Empty directories become `None` at `collect_caller_fonts` (`rdocx-py` `document.rs:17209-17272`). Diagnostic Rust methods call that same normal engine even when the slice is empty (`document.rs:46421`, `:46434`, `:46446`, `:46459`). Current `RenderOptions` at `document.rs:75-80` has `revision_view` and `revision_style` only.

Q-S96-07 asked which tier is the default. The user answered it. Alternative A is the operative policy.

## Spec reference

- `docs/hld/08-rendering-spec.md`, font caches and caller priority, lines 1440-1474, and the layout engine sections that own indent, borders, row split, colour, and symbol mapping.
- `docs/hld/10-bindings-spec.md`, native layout cache at 3191-3216, and WASM `toPdf` at 3909-3920.
- `docs/hld/15-build-and-toolchain.md`, deterministic baselines at 64-76.
- ECMA-376 Part 1 17.3.1.12 for the first-line and hanging pair. 17.7.2 and 17.4.39 for table borders.
- `docs/hld/14-development-backlog.md`, F-X278. Size L. Depends on none. Test gate differential.
- Issue 433 cases 1 through 7. Donor PR 440 is evidence of a proposed patch, not acceptance.

Section names and the backlog entry are the contract. Unrevalidated line numbers are a frozen-packet snapshot and can drift.

## Approach

All changes stay in existing files. No new source file, module, crate, trait, generic parameter, feature flag, or test binary is proposed.

Contract closure for this draft: each operation below names its callable and its A-ID. Rust, Python, and CLI routes in the tables are the single definition. Later sections refer to those tables. Full revalidation of every frozen line citation remains pending. The spot checks in Evidence pins did not invalidate this contract.

### A. TOC entry ink (A278-1)

Add two private helpers in `crates/rdocx-layout/src/engine.rs`.

`fn is_toc_paragraph_style(style_id: &str, styles: &CT_Styles) -> bool` is local to the layout crate. It returns true when the id is `TOC1` through `TOC9`, or when that style's name trims to `toc 1` through `toc 9` under ASCII case folding. `TOC 1` matches because it folds to `toc 1`. It does not call `toc_entry_style_ids`. That function is private in `crates/rdocx/src/field.rs:12209` and also needs the rebuild map. `rdocx-layout` depends on `rdocx-oxml`, not `rdocx`, and must not gain that edge.

`fn is_hyperlink_character_style(style_id: &str, styles: &CT_Styles) -> bool` matches the Hyperlink character style by id or name.

In `layout_paragraph_with_source_and_table`, `toc_entry` is a TOC paragraph style plus a hyperlink whose anchor starts with `_Toc`. Filter the Hyperlink character style from the run style at `engine.rs:11822-11829` and from the stored segment style at `:12331-12336` only when `toc_entry` is true. The body path keeps Hyperlink colour and underline. A hand-typed `_Toc` link in a TOC-named paragraph loses the Hyperlink look. Residual R-A278-1. A producer id that is neither `TOC1` through `TOC9` nor a `toc N` name stays outside this rule.

Tests use `TOC1`, `TOC 1`, `toc 1`, a paragraph style whose name is `toc 1` and whose id is not canonical, and a body Hyperlink control.

### B. First-line indent (A278-2)

Treat `w:firstLine` and `w:hanging` as one property at every `CT_PPr` merge. In `CT_PPr::merge_from` and in `merge_direct_ppr`, if either source field is present, assign both fields from the source. The absent sibling becomes `None`.

Named consumers of that paragraph merge:

- Style inheritance at `style_resolver.rs:480-527`.
- Numbering level properties at `engine.rs:11310-11319` and direct paragraph properties there.
- The `tbRl` vertical-text arm at `engine.rs:4787-4806`. Numbering level properties merge at `:4802`. Direct paragraph properties merge at `:4805`. Both calls are `merge_direct_ppr`.
- Public `Document::effective_paragraph_properties` at `document.rs:40250-40257`.
- EPUB `effective_epub_list_properties` at `epub.rs:2846-2853`.
- ODT paragraph `effective_paragraph_properties` at `odt.rs:1105-1111`. The run merge at `odt.rs:1195` is `CT_RPr::merge_from` and stays unchanged. A run-only ODT projection does not move when only `w:ind` changes.
- Style read at `style.rs:991-1005`.
- Style update merge at `document.rs:21575` and conditional table style merge at `:21656`.
- Table style paragraph merges at `table.rs:2008`, `:2052`, `:2156`, `:2162`, `:2234`, and `:2240`.
- Layout paragraph merges at `engine.rs:2061` and `:15014`.

This changes public read and import projections as well as layout. A more specific layer that sets one of the pair clears the other, matching ECMA-376 Part 1 17.3.1.12. Tests cover the public, EPUB, and ODT paragraph consumers. They also cover one `tbRl` paragraph whose earlier layer sets hanging and whose later layer sets only firstLine, and the swapped pair. The inherited value is the one that remains. The stale sibling is `None`. The assertion uses the existing twip positions. Unit rounding stays as it is.

A single `w:ind` that carries both attributes stays in the model and in raw XML. `line.rs:2256-2260` and `:2307-2310` keep hanging precedence. That is residual R-A278-2 and native calibration debt. No unit rounding change.

### C. Table borders (A278-3)

Stop promoting the base style's `w:tblBorders` into cell borders in `resolve_table_style_cell` at `table.rs:2165-2172`. Style table borders still reach the table through `resolve_base_table_properties` at `:1054-1135`, under direct table properties. Paint prefers a cell edge and falls back to the table edge at `paginator.rs:7817-7945`. Because `resolved_cell_edge` treats cell `none` as unset at `table.rs:983-992`, copying `none` onto cells is not the fix.

Acceptance is only the issue pair. Style-only TableGrid is 3 horizontal and 3 vertical ruled lines at 72 dpi. Direct `w:tblBorders` with all six edges `none` is 0 and 0. Partial edges and conditional promotion at `table.rs:2245-2253` stay residual R-A278-3 with no count.

The donor `dense_form` numbers are not accepted. Re-derive that local expectation after the union. It is not a hash key.

### D. Row split (A278-4)

Remove only the `cell.v_align.is_some()` bail at `paginator.rs:1844`. Keep `cant_split`, header rows, merges, clipping, rotation, anchored content, keep-next, and page-break-before as whole-row barriers. A short cell beside a long one is aligned inside the first split part. Word placement of that short cell is residual R-A278-4.

Discriminator: the issue row starts on page 1 and the first-page line count stays within one line of base `[23, 17]`.

### E. Automatic text colour (A278-5)

Add `fn automatic_text_color(fill: Option<Color>) -> Color` in `engine.rs`. No fill returns black. Otherwise round each channel from the 0.0 to 1.0 range, clamp, and scale to 0 through 255. Then `floor((76 * r + 151 * g + 29 * b) / 256) <= 156` selects white. The coefficients and the floor are the donor's chosen policy. They are not a pinned LibreOffice source file. The three issue colours are the acceptance. Word's threshold stays calibration debt inside R-A278-5.

Thread `cell_shading: Option<Color>` through these existing functions, placed before `doc_grid`:

- `layout_cell_content`
- `layout_control_cell_content`
- `layout_paragraph_with_source_in_table`
- `layout_paragraph_with_source_and_table`

Call sites:

- `layout_table_inner` computes shading at `table.rs:690-697` and passes it into `layout_cell_content` at `:755`.
- `autofit_column_widths` passes `None` at the measurement call `:1510`. Colour does not change the measured width.
- The paragraph arm inside `layout_cell_content` at `table.rs:1735` forwards that same `cell_shading` into `layout_paragraph_with_source_in_table`. It does not pass `None` and it does not recompute shading.
- Nested `CellContent::Table` at `:1771` calls `layout_table_inner` and does not receive the outer cell's shading. The nested cell uses only the shading that inner computes at `:690-697`.
- `layout_cell_content` forwards the same `Option` into `layout_control_cell_content` at `:1789`.
- The control paragraph arm at `table.rs:1869` forwards the `cell_shading` received by `layout_control_cell_content` into `layout_paragraph_with_source_in_table`. A content control in a shaded cell uses that forwarded fill.
- The recursive content-control call at `:1923` forwards it again.
- `layout_paragraph_with_grid` at `engine.rs:11149` passes `None`.
- `layout_paragraph_with_source` at `engine.rs:11180` passes `None`.
- `layout_paragraph_with_source_and_direction` at `engine.rs:11215` passes `None` into `layout_paragraph_with_source_and_table`. This wrapper is not a cell. A body paragraph laid out through it keeps black when it has no paragraph shading.

The control-paragraph test uses a shaded cell whose content is a content control with one paragraph and automatic colour. The nested-table control uses an outer shaded cell and an inner unshaded cell. The inner automatic run stays black. The body control through `layout_paragraph_with_source_and_direction` stays black.

`automatic_color` is the paragraph's own shading, or else the cell fill. Pass it through `resolve_run_color`, the numbering marker, annotation lines, tracked marks, equations, carets, and segment colours. An explicit `w:color` wins. Run shading and highlighting are not the background. Residual R-A278-5.

### F. Wingdings 2 (A278-6)

`oxml-layout` gains exactly three `#[doc(hidden)]` methods on `FontManager`. They are additive. `current_font_diagnostics` at `font.rs:1850-1852` stays a separate method and its body stays `&self.layout_font_diagnostics`.

```rust
#[doc(hidden)]
pub fn resolve_word_symbol_text(
    &mut self,
    family: Option<&str>,
    bold: bool,
    italic: bool,
    original: &str,
    ppt_bullet_compatibility: bool,
) -> Result<(String, FontId)>

#[doc(hidden)]
pub fn unknown_word_symbol_diagnostics(
    &self,
    selected: FontId,
    requested: Option<&str>,
    original: &str,
    displayed: &str,
) -> Vec<String>

#[doc(hidden)]
pub fn word_substitution_diagnostics(&self) -> Vec<Diagnostic>
```

`resolve_word_symbol_text` and `unknown_word_symbol_diagnostics` share a private helper with the existing methods. Only the Word methods add the Wingdings 2 table. Existing behaviour of `resolve_symbol_text` and `unknown_symbol_diagnostics` stays unchanged when that private helper is shared.

Switch only these `rdocx-layout` calls to the Word methods: `engine.rs:11453`, `:11480`, `:12576`, and `:12590`. Leave these PowerPoint calls on the existing methods:

- `rpptx-render/src/text.rs:515` `resolve_symbol_text`
- `rpptx-render/src/text.rs:912` `unknown_symbol_diagnostics`
- `rpptx-render/src/text.rs:3729` `resolve_symbol_text`
- `rpptx-render/src/text.rs:3752` `resolve_symbol_text`
- `rpptx-render/src/text.rs:3710` `current_font_diagnostics`
- `rpptx-render/src/lib.rs:367` `current_font_diagnostics`

`word_substitution_diagnostics` reads a separate vec. Shared font resolution may append a fact to that vec when a requested family is absent and a fallback face is selected. It must not push that fact into `layout_font_diagnostics`. `rdocx-layout` appends the new method at the existing aggregation points `engine.rs:4318` and `:6338`, beside the unchanged `current_font_diagnostics` extend. PowerPoint never calls the new method, so its diagnostic list stays today's list.

Each substitution fact names the requested family, the selected physical family, and the tier `caller`, `embedded`, `system`, or `bundled`. The Aptos approximation message stays on the old method. CLI `convert` prints diagnostics at `commands.rs:836-837`. Python `to_pdf_with_diagnostics` returns them. `page_count`, `layout`, and `layout_page` still discard them. Residual R-A278-7b. The message contains the requested family and the selected physical family, so a DejaVu Sans fallback can satisfy the issue substring `DejaVu Sans`. No such message is claimed as an executed result. The PDF `BaseFont` check remains the other script channel.

Bounded Wingdings 2 table, low byte and `U+F000` form: `0x4F` to `U+2717`, `0x50` to `U+2713`, `0x52` to `U+2611`, `0x54` to `U+2612`, `0x96` and `0x98` to `U+25CF`, `0x97` to `U+2022`, `0x99` through `0x9C` to `U+25CB`, `0x9F` and `0xA0` to `U+25AA`, `0xA1` and `0xA2` to `U+25A0`, `0xA3` to `U+25A1`, `0xAD` to `U+25C6`, `0xAE` to `U+2B25`, `0xB7` to `U+2B27`. `U+F0A1` becomes `U+25A0`. Unmapped `U+F0A7` keeps its scalar and reports `Unknown Wingdings 2 symbol U+F0A7: original code retained`. Genuine coverage still wins through `has_original_symbol_coverage`. This is not measured glyph coverage. No new bundled face or licence. The 32 faces in `bundled_fonts.rs` stay. PowerPoint has its own F0B7 path and no Issue 433 acceptance, so the shared function must not gain this table. A future PowerPoint Wingdings 2 story is separate.

### G. Shared font policy (A278-7)

One private owner in `crates/rdocx/src/document.rs`:

```rust
enum FontPolicy {
    BundledOnly,
    SystemWithBundledFallback,
}

fn layout_with_font_policy(
    &self,
    policy: FontPolicy,
    font_files: &[(&str, &[u8])],
    options: RenderOptions,
) -> Result<Arc<rdocx_layout::WordLayoutResult>>
```

The return is always `Result<Arc<WordLayoutResult>>`. One signature covers the cache and the owned caller-font path.

- Empty `font_files` and an accepted revision view return the cache `Arc` unchanged. `BundledOnly` returns `cached_deterministic_layout()` (`document.rs:25020`). `SystemWithBundledFallback` returns `cached_layout()` (`document.rs:24995`). A cache hit inside those methods is `Arc::clone` at `:25026` and `:25001`. `Arc::clone` preserves sharing. Cloning the inner `WordLayoutResult` is rejected because that would break the shared result.
- Empty `font_files` and a tracked revision view stay uncached. The helper builds an owned result and returns `Arc::new` of it. That `Arc` is not stored in either cache.
- Nonempty `font_files` call the existing owned public method for the policy and return `Arc::new` of that owned result. `BundledOnly` calls `layout_with_fonts_and_bundled_fallback_and_options`. `SystemWithBundledFallback` calls `layout_with_additional_fonts_and_options`. This `Arc` is not the omission cache. `Arc::new` does not clone a cached allocation. The explicit methods keep their owned `Result<WordLayoutResult>` returns.

`layout` and `layout_with_options` already return `Result<Arc<WordLayoutResult>>`. They return the helper `Arc` unchanged. Priority inside either tier is caller, then document-embedded, then the selected tier. Caller priority stays highest. The selected omission tier is bundled. System is only the explicit opt-in.

The helper is private to `rdocx`. `rdocx-py` and `rdocx-cli` cannot call it. Their nonempty routes below are existing public methods. Setting `use_system_fonts` on a method in the ignore list does not select the tier.

`FontManager::new`, `new_deterministic`, and the `system-fonts` compile paths stay as they are. `new` loads system fonts only inside `#[cfg(feature = "system-fonts")]` at `font.rs:1036-1046` and loads the bundled database at `:1049-1050` otherwise. `new_deterministic` is always bundled.

#### Empty input is omission

`None`, an empty slice, `fonts=[]`, an empty directory, and a directory with no font file are omission. They are not a caller-font set. `collect_caller_fonts` already returns `None` for an empty selection. Python diagnostic methods today replace that `None` with `&[]` and then call the normal additional-font engine. That replacement is the bug. After the change, empty input is omission and uses the bundled tier unless `use_system_fonts`, `system_fonts`, or `--system-fonts` explicitly selects system. Inside `rdocx`, the diagnostic methods call the private helper for that selected tier. Python and CLI do not call the helper. They use the public route tables below. A nonempty list keeps caller priority and does not switch the tier.

#### The Rust control

`RenderOptions` at `document.rs:75-80` gains one public field:

```rust
pub use_system_fonts: bool
```

The struct stays `Copy`, `PartialEq`, `Eq`, and derived `Default`. `use_system_fonts` therefore defaults to false. No manual `Default` is added.

Adding the field breaks literals that name `revision_view` and `revision_style` and do not use `..`. The CLI literal at `commands.rs:2248-2253` is one of those sites. Every such literal sets the field from the flag or from `..RenderOptions::default()`. The field on an ignore-list call is not the tier switch. The call selected below is the switch.

This field is the minimal callable control. The implicit Rust publishers already take `RenderOptions` or call a sibling that does. A new public selector method is not required. A policy enum in `oxml-layout` is rejected because it would sit on the shared crate. Without this field, there is no zero-font Rust call that can request the system opt-in. Selecting Alternative A does not change the legacy behaviour of caller-only, bundled-fallback, alias, additional-fonts, or deterministic methods. Those methods stay on the ignore list.

Readers of the field are the implicit publishers in the inventory. These explicit methods ignore the field and keep today's engine:

- `layout_deterministic` and `layout_deterministic_with_options` (`:45824-45836`)
- `to_pdf_deterministic`, `to_pdf_deterministic_with_options`, `to_pdfa_deterministic` (`:46070-46102`)
- `to_pdfa_with_diagnostics` (`:46084-46091`), which stays bundled-fallback
- `to_pdf_with_fonts_deterministic_and_options` and `to_pdfa_with_fonts_deterministic`
- `render_page_to_png_deterministic`, `render_page_to_svg_deterministic`, `render_pages_deterministic`, and their `_with_options` forms
- `layout_with_fonts` and `layout_with_fonts_and_options`, caller-only
- `layout_with_fonts_and_bundled_fallback` and its options form, caller plus bundled
- `layout_with_fonts_aliases_and_bundled_fallback` (`:45983`) and `layout_with_fonts_aliases_and_bundled_fallback_and_options` (`:45996`), caller aliases plus bundled
- `layout_with_additional_fonts_and_options` (`:46107-46115`), caller plus the normal engine
- `to_pdf_with_fonts` and `to_pdf_with_fonts_and_options` (`:46154-46169`)
- `render_page_to_png_with_fonts_and_options`, `render_page_to_svg_with_fonts_and_options`, `render_pages_with_fonts_and_options`, and `render_all_pages_with_fonts_and_options`

Deterministic helpers ignore `use_system_fonts` so a caller cannot move the hash path. The harness stays on those helpers. `layout_deterministic` at `document.rs:45824` was re-read and still exists.

#### Implicit Rust publishers

Each row's no-caller call today is the normal engine unless the current column says bundled. Omission is bundled because `use_system_fonts` defaults to false. Explicit true is the system opt-in. Diagnostic and non-diagnostic siblings share that choice. The system column below is that opt-in.

| Publisher | Today, no caller font | Empty-slice rule after the change |
|---|---|---|
| `layout` `:45805`, `layout_with_options` `:45813` | normal, `layout_for_options(..., false)` | read the field |
| `layout_page` `:46529`, `layout_page_with_options` `:46534` | normal, calls `layout_with_options` at `:46539` | read the field |
| `to_pdf` `:46055`, `to_pdf_with_options` `:46060` | normal | read the field |
| `save_pdf` `:46173`, `save_pdf_with_options` `:46178` | normal, calls `to_pdf_with_options` at `:46183` | read the field |
| `render_page_to_png` `:46260`, `render_page_to_png_with_options` `:46265` | normal, `layout_with_options` at `:46271` | read the field |
| `render_page_to_svg` `:46316`, `render_page_to_svg_with_options` `:46321` | normal, `layout_with_options` at `:46326` | read the field |
| `render_pages` `:46349`, `render_pages_with_options` `:46358` | normal, `layout_with_options` at `:46364` | read the field |
| `render_all_pages` `:46401`, `render_all_pages_with_options` `:46406` | normal, `layout_with_options` at `:46411` | read the field |
| `to_pdf_with_diagnostics` `:46416` | always `layout_with_additional_fonts_and_options` at `:46421`, even for `&[]` | reads the field for empty and nonempty. Empty returns the sibling cache `Arc`. Nonempty calls the policy's owned public layout method and `Arc::new` |
| `render_page_to_png_with_diagnostics` `:46427` | always additional-fonts at `:46434` | same helper rule as `to_pdf_with_diagnostics` |
| `render_all_pages_with_diagnostics` `:46440` | always additional-fonts at `:46446` | same helper rule as `to_pdf_with_diagnostics` |
| `render_pages_with_diagnostics` `:46452` | always additional-fonts at `:46459` | same helper rule as `to_pdf_with_diagnostics` |

`layout()`, `to_pdf()`, `save_pdf()`, `layout_page()`, and the render methods that take no `RenderOptions` call the options sibling with `RenderOptions::default()`. Their opt-in is that sibling. There is no second flag.

#### Implicit Python publishers

Today, no-caller `page_count`, `layout`, and `layout_page` call `layout_deterministic` (`:12510`, `:15783`, `:15832`). No-caller `to_pdf` and the render helpers call the normal engine. Diagnostic helpers always pass `&[]` into the additional-fonts Rust method (`to_pdf_with_diagnostics` at `:13538-13540`, PNG diagnostics at `:13619-13623`, `render_all_pages_with_diagnostics` at `:13719-13723`, `render_pages_with_diagnostics` at `:13822-13826`). `render_all_pages` at `:13689-13693` already branches: `None` calls `render_all_pages_with_options`, and a font list calls the fonts method.

After the change, `None` and an empty slice are omission and use the row's default tier through the public omission call in the route table. A nonempty list keeps caller priority and uses the public nonempty call for that same tier. It does not call the private helper. It does not call a normal-engine method when the tier is bundled. Every method in this list takes the alternative's keyword. The keyword selects the row in the route table. R-A278-7c stays visible.

| Python method | Today, no caller font | Keyword |
|---|---|---|
| `page_count` `:12494` | bundled | same keyword as `layout` |
| `layout` `:15765` | bundled | same keyword |
| `layout_page` `:15815` | bundled, `layout_deterministic` at `:15832` | same keyword, same default, same error |
| `to_pdf` | normal | same keyword |
| `to_pdf_with_diagnostics` `:13521` | normal through empty-slice additional-fonts | same keyword |
| `render_page_to_png` | normal | same keyword |
| `render_page_to_png_with_diagnostics` `:13600` | normal through empty-slice additional-fonts | same keyword |
| `render_page_to_svg` `:13630` | normal | same keyword |
| `render_pages` `:13734` | normal | same keyword |
| `render_pages_with_diagnostics` `:13800` | normal through empty-slice additional-fonts | same keyword |
| `render_all_pages` `:13670` | normal, `render_all_pages_with_options` | same keyword |
| `render_all_pages_with_diagnostics` `:13701` | normal through empty-slice additional-fonts | same keyword |

`to_pdfa_deterministic` at `:13546` stays bundled, takes no font-policy keyword, and is not an issue entry point.

Python omission calls for the selected policy:

| Python method | Bundled omission | System omission |
|---|---|---|
| `page_count`, `layout`, `layout_page` | `layout_deterministic` | `layout_with_options` with `use_system_fonts: true` |
| `to_pdf` | `to_pdf_deterministic_with_options` | `to_pdf_with_options` with true |
| `to_pdf_with_diagnostics` | `to_pdf_with_diagnostics` with false and an empty slice | the same method with true and an empty slice |
| `render_page_to_png` | `render_page_to_png_deterministic_with_options` | `render_page_to_png_with_options` with true |
| `render_page_to_png_with_diagnostics` | `render_page_to_png_with_diagnostics` with false and an empty slice | the same method with true and an empty slice |
| `render_page_to_svg` | `render_page_to_svg_deterministic` | `render_page_to_svg_with_options` with true |
| `render_pages` | `render_pages_deterministic_with_options` | `render_pages_with_options` with true |
| `render_pages_with_diagnostics` | `render_pages_with_diagnostics` with false and an empty slice | the same method with true and an empty slice |
| `render_all_pages` | `layout_deterministic_with_options`, then `oxml_pdf::render_all_pages` | `render_all_pages_with_options` with true |
| `render_all_pages_with_diagnostics` | `render_all_pages_with_diagnostics` with false and an empty slice | the same method with true and an empty slice |

Bare `layout()` and bare `to_pdf()` use `RenderOptions::default()`. That default is false, so they are bundled omission and not the system opt-in. The system column passes true explicitly. Revision fields from the existing parser are copied onto that struct.

Bundled omission for `render_all_pages` calls `layout_deterministic_with_options` and then the existing public `oxml_pdf::render_all_pages`. That paint is already public at `oxml-pdf/src/lib.rs:63`. `rdocx-py` already calls `oxml_pdf::render_pages` at `rdocx-py/src/document.rs:13767`. This is the same dependency, not a new crate edge. System omission calls `render_all_pages_with_options`, which reads the field.

The four diagnostic methods are public policy consumers. They read `use_system_fonts` for an empty slice and for a nonempty slice. They are not on the ignore list. `use_system_fonts` true with `system-fonts` off returns the named error before either engine runs, for an empty slice and for a nonempty slice. There is no silent bundled success for that explicit request.

Python nonempty calls. The bundled column is omission, `system_fonts` false. The system column is the explicit opt-in, `system_fonts` true, on a build that has `system-fonts`. Python does not grow a `bundled_fonts` keyword.

| Python siblings | Bundled nonempty public call | System nonempty public call |
|---|---|---|
| `page_count`, `layout`, `layout_page` | `layout_with_fonts_and_bundled_fallback_and_options` | `layout_with_additional_fonts_and_options` |
| `to_pdf` and `to_pdf_with_diagnostics` | `layout_with_fonts_and_bundled_fallback_and_options`, then public `oxml_pdf::render_to_pdf_checked` mapped with `rdocx::Error::from` | `to_pdf_with_diagnostics` |
| `render_page_to_png` and `render_page_to_png_with_diagnostics` | the bundled layout call, then public `oxml_pdf::render_page_to_png` | `render_page_to_png_with_diagnostics` |
| `render_page_to_svg` | the bundled layout call, then `rdocx::render_page` | `render_page_to_svg_with_fonts_and_options` |
| `render_pages` and `render_pages_with_diagnostics` | the bundled layout call, then public `oxml_pdf::render_pages` | `render_pages_with_diagnostics` |
| `render_all_pages` and `render_all_pages_with_diagnostics` | the bundled layout call, then public `oxml_pdf::render_all_pages` | `render_all_pages_with_diagnostics` |

A system-column call that reads the field passes `use_system_fonts: true`. A bundled diagnostic call that reads the field passes false. `layout_with_additional_fonts_and_options` ignores the field, so the function name is the system tier, not the bool. `to_pdf_with_fonts_and_options`, `render_page_to_png_with_fonts_and_options`, `render_page_to_svg_with_fonts_and_options`, `render_pages_with_fonts_and_options`, and `render_all_pages_with_fonts_and_options` stay on the normal engine and stay on the ignore list. Python calls them only where the route table names `render_page_to_svg_with_fonts_and_options`. The other four normal-engine font methods are not the policy switch. A nonempty bundled request must not call them.

On a system nonempty route, the non-diagnostic sibling calls the same public diagnostic method as its twin and discards the message tuple. Both therefore use `layout_with_additional_fonts_and_options` inside that method. `page_count`, `layout`, `layout_page`, and `render_page_to_svg` have no diagnostic twin. They call the public layout or SVG method named in the system column.

Alternative A feature-off gate for a system nonempty request: call `layout_with_options` with `use_system_fonts: true` first when the sibling is `page_count`, `layout`, `layout_page`, or `render_page_to_svg`. Return that error and do not call the normal font method. `layout_with_additional_fonts_and_options` would succeed on a feature-off build because `FontManager::new` loads the bundled database, so it is not the error route. The diagnostic methods already return the error themselves, so PDF, PNG, `render_pages`, and `render_all_pages` do not add this probe. An accepted-view probe can fill the normal omission cache. A tracked-view probe does not. Tests that count layout invocations on those four methods must expect the probe. The probe is the feature-off refusal for those four methods, including a nonempty caller font.

`layout` and `layout_page` today wrap a nonempty owned result in a local `Arc::new` at `rdocx-py/src/document.rs:15785` and `:15834`. That local `Arc` is not the shared cache. Omission keeps the `Arc` returned by `layout_with_options` or `layout_deterministic`. Do not unwrap a cached `Arc` and do not clone the inner `WordLayoutResult`.

SVG paint is `svg::render_page` at `crates/rdocx/src/svg.rs:29`. It is `pub(crate)` and `mod svg` is private at `lib.rs:44`. `rdocx-py` cannot call it. `render_page_to_svg_with_fonts_and_options` at `document.rs:46489` always uses the normal engine, so it is the wrong bundled call. Closure: change `render_page` from `pub(crate)` to `pub` and add `pub use svg::render_page` in `lib.rs`. No new function, file, trait, or selector. The normal SVG method stays unchanged and stays on the ignore list. This visibility change is a Word 0.17.0 public item on the first unpublished Word line. It is not one of the three shared 0.15.0 additions and it is not an oxml-layout 0.16.0 bump.

One caller font is the discriminator. For bundled omission, `to_pdf(font_dir=DIR)` and `to_pdf_with_diagnostics(font_dir=DIR)` both use the bundled layout call and the same page count. For explicit system opt-in, that same pair both use `to_pdf_with_diagnostics` and the same page count. The same pair rule applies to PNG, `render_pages`, `render_all_pages`, `layout`, and `layout_page`.

#### CLI

| Command | Today, no caller font | Flag |
|---|---|---|
| `convert` | bundled fallback, load `:811-817`, PDF layout `:833`, image layout `:857`, print `:836-837` | `--system-fonts`, default off |
| `render` | bundled fallback at `:2248-2253` | the same flag, default off |
| `layout` | bundled fallback at `:438` | the same flag, default off. It still does not print diagnostics |
| `inspect` | `layout_deterministic` at `:106` and `:141`. The text label is "Pages (deterministic layout)" | no flag. Stays on `layout_deterministic` |

`inspect` is outside the four issue entry points. With no `--system-fonts`, `convert` omission and `inspect` both use the deterministic tier. With `--system-fonts`, `convert` uses the system opt-in and `inspect` stays deterministic, so those page counts may differ. CLI does not grow `--bundled-fonts`.

Empty `font_refs` are omission. The current calls pass that empty slice into bundled-fallback methods, which ignore `use_system_fonts`. After the change those four sites branch. HTML and Markdown stay without a font flag, as they do today at `commands.rs:808-809`.

| Site | Omission, bundled tier | Omission, system tier | Nonempty, bundled tier | Nonempty, system tier |
|---|---|---|---|---|
| `convert` PDF `:833` and images `:857` | `layout_deterministic_with_options` | `layout_with_options` with `use_system_fonts: true` | `layout_with_fonts_and_bundled_fallback_and_options` | `layout_with_additional_fonts_and_options` |
| `render` `:2248` | `layout_deterministic_with_options` | `layout_with_options` with true | the bundled-fallback options method | `layout_with_additional_fonts_and_options` |
| `layout` `:438` | `layout_deterministic` | `layout_with_options` with `use_system_fonts: true` | `layout_with_fonts_and_bundled_fallback` | `layout_with_additional_fonts_and_options` |

Paint stays the existing `oxml_pdf::render_to_pdf_checked` at `:834` and `oxml_pdf::render_pages` at `:861` and in `render`. The branch replaces only the layout call. Diagnostic print at `:836-837` still reads `layout.layout.diagnostics` from whichever layout call returned.

`--system-fonts` on a feature-off CLI build calls `layout_with_options` with true and returns the named error. That call is the gate for an empty font list and for a nonempty list. A nonempty system request does not continue to `layout_with_additional_fonts_and_options` after the error. No flag on a feature-off build stays on the bundled omission column and does not return that error.

`rdocx-cli` default features include `system-fonts` at `crates/rdocx-cli/Cargo.toml:30-31`. `cargo test -p rdocx-cli --no-default-features` is the CLI feature-off invocation. It is not a substitute for the `rdocx` lib test below.

Discriminators, one caller font and the no-font command, local and final:

- `rdocx convert --system-fonts` selects the system column. `rdocx convert` selects the bundled column. The same pair holds for `rdocx render` and `rdocx layout`.
- The accepted `rdocx inspect` invocation has no `--system-fonts` flag and stays on `layout_deterministic`. Passing `--system-fonts` to `inspect` is an unsupported-flag refusal. Inspect does not gain that option.

#### Selected policy, Alternative A

- Rust keeps derived `Default`. `RenderOptions::default().use_system_fonts` is `false`.
- Python keyword-only `system_fonts: bool = False` on every method in the Python table.
- CLI `--system-fonts`, default off, on `convert`, `render`, and `layout`.
- Omission uses the bundled column, which is the deterministic cache for an empty font list. `true` or the flag selects the system column, then bundled faces for families the system font set does not contain.
- Feature off: `use_system_fonts: true`, `system_fonts=True`, and `--system-fonts` return `rdocx::Error::Other("system fonts are not available in this build (the system-fonts feature is disabled)")` before the normal engine runs. This applies to empty and nonempty caller fonts. Python maps that error through `rdocx_to_pyerr` at `rdocx-py/src/lib.rs:131` when the linked `rdocx` was built without the feature. There is no silent downgrade of an explicit system request.
- `rdocx-py/Cargo.toml:34` always enables `rdocx` features `system-fonts` and `digital-signatures`. `cargo test -p rdocx-py --no-default-features` does not turn that dependency feature off and is not the feature-off evidence. The proving command is `cargo test -p rdocx --no-default-features --lib` in the existing document test module. `cargo test -p oxml-layout --no-default-features` remains the shared bundled-database rider only.
- WASM exports no font-policy option. `toPdf` calls `self.inner.to_pdf()` at `rdocx-wasm/src/lib.rs:85`. Under this alternative that call uses the deterministic cache, which matches desktop omission. `rdocx-wasm` sets `default-features = false` at `crates/rdocx-wasm/Cargo.toml:27` and does not enable `system-fonts`. A future wasm system request must refuse with the same error. No binding route change.
- Changed versus today: Rust implicit publishers and Python `to_pdf`, `to_pdf_with_diagnostics`, and every render helper move from normal to bundled on omission. CLI omission and Python `page_count`, `layout`, and `layout_page` are already bundled and stay on the deterministic cache. Nonempty Python and CLI calls move to the route table.
- Explicit deterministic methods stay bundled. Explicit caller-only, bundled-fallback, alias bundled-fallback, and additional-fonts Rust methods stay on their documented engines and ignore the field.

#### Alternative B is not selected

The unselected draft used a manual `Default` with `use_system_fonts` true, a Python `bundled_fonts` keyword, and a CLI `--bundled-fonts` flag. Its feature-off omission was bundled normal-engine output through `cached_layout`, and its wasm `toPdf` was bundled normal-engine output while desktop omission could load system fonts. Those clauses are not applicable. The system opt-in through `use_system_fonts` true, Python `system_fonts=True`, and CLI `--system-fonts` stays in Alternative A. Python does not gain `bundled_fonts`. CLI does not gain `--bundled-fonts`.

The selected policy keeps caller fonts highest. Empty input is omission. `font_dir` and `font_dirs` stay mutually exclusive. Word `rdocx`, `rdocx-cli`, and `rdocx-py` stay on the first unpublished 0.17.0 line from F-X275. The `RenderOptions` field, the Python `system_fonts` keyword, the CLI `--system-fonts` flag, and the `render_page` visibility change are Word 0.17.0 public changes. No 0.16.x patch break.

`.claude/plans/F-X275-design.md:31` prepares the shared family at 0.15.0. Line 49 identifies 0.15.0 and 0.17.0 as the first unpublished minor lines. `publish = true` at `crates/oxml-layout/Cargo.toml:12` proves that the crate may be published. It does not prove that 0.15.0 has been released. The three `#[doc(hidden)]` `FontManager` methods are public additions. They stay in the shared 0.15.0 SemVer inventory and in the exact published-baseline and compiler checks for that first unpublished line. They do not require an oxml-layout 0.16.0 bump. The shared public API is changed by those three methods. Do not describe it as unchanged.

PowerPoint call sites and PowerPoint symbol output stay unchanged. That is a separate claim from the shared API inventory. `rpptx` and `rpptx-py` behaviour stays unchanged. No publication happens in this preparation.

## Consumer map

| Rule | Entry points and outputs | Shared provider and caller policy | Lifecycle stages compared |
|---|---|---|---|
| TOC entry ink | `layout_paragraph_with_source_and_table`, run style `engine.rs:11822-11829`, segment style `:12331-12336` | Local `TOC1` through `TOC9` and `toc N` names, `_Toc` anchor | Parse, layout, render, save and reopen |
| First-line indent | `CT_PPr::merge_from`, `merge_direct_ppr`, public, EPUB, ODT paragraph, `tbRl` at `engine.rs:4802` and `:4805`, `line.rs` | One first-line property. Same-element both attributes preserved | Parse, style resolve, public read, import, layout, authoring reopen |
| Table borders | `resolve_base_table_properties`, `resolve_table_style_cell`, paint edges | Direct table borders over style table borders. Cell `none` is unset | Parse, style resolve, layout, paginate, render |
| Row split | `split_simple_table_row` | `w:vAlign` does not gate splitting. `w:cantSplit` still does | Layout, paginate, render |
| Automatic colour | `resolve_run_color`, cell fill at `table.rs:755`, `:1510`, `:1735`, `:1771`, `:1789`, `:1869`, `:1923`, and `None` at `engine.rs:11149`, `:11180`, `:11215` | Paragraph shading, else this cell's fill. Nested cell does not inherit. Body wrapper passes `None` | Parse, layout, render |
| Wingdings 2 | Three Word-only `FontManager` methods, four `engine.rs` sites | Bounded map. PowerPoint sites listed above stay on the old methods | Parse, layout, render, diagnostics, reopen |
| Font policy | Rust, Python, and CLI route tables, including `convert`, `render`, and `layout` | Caller over embedded over the selected tier. Empty input is omission. Siblings share one public route. `inspect` stays deterministic | Open, held and fresh layout, render, save and reopen |

## Behavior contracts and evidence

| Surface or placement | Exact behavior and boundary examples | Primary source identity and location | Local discriminator or pending evidence |
|---|---|---|---|
| TOC ink | TOC paragraph plus `_Toc` draws `#000000` and no underline. Body Hyperlink draws `#0563C1` and underline | Issue 433 lines 13-78, `field.rs:10395-10420` | `TOC1`, `TOC 1`, `toc 1`, non-canonical id with name `toc 1`, body control. Word and LibreOffice pending |
| Indent | Either `w:firstLine` or `w:hanging` clears the other. Four issue cases start at 72.0 pt, tolerance 0.05 pt. Both attributes on one `w:ind` keep hanging precedence. `tbRl` uses the same merge | ECMA-376 17.3.1.12, `line.rs:2256-2260` and `:2307-2310`, `engine.rs:4802` and `:4805` | Four inherited cases, public, EPUB, ODT, and one `tbRl` pair. Native calibration pending |
| Table borders | Style-only TableGrid is 3 and 3. Direct all-`none` is 0 and 0 | ECMA-376 17.7.2 and 17.4.39, `table.rs:983-992` | 72 dpi PNG. Partial edges are R-A278-3 |
| Row split | Top, center, and bottom `w:vAlign` split with the unaligned row. Page 1 stays within one line of `[23, 17]` | Issue 433 lines 226-274, `paginator.rs:1844` | Three alignments plus base. LibreOffice center difference is R-A278-4 |
| Automatic colour | Black `000000` and navy `1F3864` draw `#FFFFFF`. White or no fill draws `#000000`. Explicit black stays black. A control paragraph in a shaded cell uses the forwarded fill. A nested unshaded cell stays black | Issue 433 lines 276-316. Coefficients are chosen policy | Three issue cells, light grey, no fill, explicit black, paragraph shading, control paragraph, nested table, body `None`. Word threshold is R-A278-5 |
| Wingdings 2 | `U+F0A1` draws `U+25A0`. Unmapped `U+F0A7` keeps its scalar plus the named diagnostic. Wingdings `U+F0A7` draws `U+25AA`. Symbol `U+F0B7` draws `U+2022` | Issue 433 lines 318-382. Chosen bounded table | Three-font control. PowerPoint output unchanged. Coverage unexecuted |
| Font policy | One font tier and one page count across CLI `convert`, `render`, and `layout`, and Python `to_pdf`, `page_count`, `layout`, and `layout_page`, for each folder. Omission is bundled. `--system-fonts` and `system_fonts=True` select system. Each diagnostic publisher uses the same public route as its sibling for a nonempty font. `inspect` stays deterministic | Issue 433 lines 384-452. Route tables above | No folder, empty folder, one caller font, flag off and flag on for the three commands. DejaVu Sans and the CLI binary are final-union facts. Q-S96-07 is answered and does not block |

## Dependency readiness

No provider edges. F-X274's approved plan and reviewed engine source are a resource base, not a formal provider. F-X272 is not a provider and F-X278 does not depend on it. Source Root final 4 rejected C1 and C4. Source work is paused for design clarification. F-X276 is a canonical draft, pending and unclaimed, and it is not a provider.

This draft is prepared from the frozen 919-file packet at head `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, kind `as-of-stopped-source-text-only-not-clean-worktree-or-approval`. The scoped current-worktree reads in Evidence pins do not replace a full citation revalidation and do not claim a clean Base. Claim and the shared approved-design commit stay on the workflow gates. This adoption does not add a clean-Base prerequisite for the draft itself.

F-X273 and F-X274 affected output evidence is re-run once in the final union. Approval of this plan does not mark those stories done.

## Acceptance criteria

| ID | Input or setup | Expected observation or refusal | Consumers | Local test | Final gate or pending evidence |
|---|---|---|---|---|---|
| A278-1 | Word-like TOC run with Hyperlink, plus an rdocx-built TOC | Both entries `#000000` and not underlined. Body Hyperlink keeps colour and underline | Run style and segment style sites | `issue433_toc_entries_drop_the_hyperlink_character_style` | Issue 433 case 1. Word and LibreOffice pending |
| A278-2 | BaseInd, derived firstLine, direct firstLine, DerivedH, same `w:ind` with both attributes, `tbRl` inherited pair | Four issue cases at 72.0 pt. Same-element pair keeps hanging position. `tbRl` clears the absent sibling | `merge_from`, `merge_direct_ppr`, public, EPUB, ODT paragraph, `engine.rs:4802` and `:4805` | `issue433_first_line_indent_is_one_property_across_levels` plus those consumers and the `tbRl` pair | Issue 433 case 2 |
| A278-3 | TableGrid only, and TableGrid plus direct all-`none` | 3 and 3, then 0 and 0 | `resolve_table_style_cell`, paint | `issue433_direct_table_borders_override_the_style_grid` | Issue 433 case 3 at 72 dpi |
| A278-4 | 15 fillers, 40-line row, no `cantSplit`, three `vAlign` values | Row starts on page 1 and stays within one line of `[23, 17]` | `split_simple_table_row` | `issue433_cell_vertical_alignment_does_not_stop_a_row_from_splitting` | Issue 433 case 4 |
| A278-5 | Auto on black, navy, and white, plus a shaded-cell content control, a nested unshaded cell, and a body paragraph | `#FFFFFF`, `#FFFFFF`, `#000000`. Control paragraph follows the cell. Nested and body controls stay black | `resolve_run_color` and every named shading call, including `:1735`, `:1869`, and `engine.rs:11215` | `issue433_automatic_text_colour_contrasts_with_the_fill_behind_it` | Issue 433 case 5. Word threshold pending |
| A278-6 | Wingdings 2 `U+F0A1`, Wingdings `U+F0A7`, Symbol `U+F0B7`, unmapped Wingdings 2 | `U+F0A1` to `U+25A0`, controls unchanged, unmapped diagnostic, PowerPoint unchanged | Word symbol methods only | `issue433_wingdings_2_bullets_map_to_unicode_or_report_the_code` | Issue 433 case 6. Coverage unexecuted |
| A278-7 | No folder, empty folder, one caller font, `--system-fonts` off and on | One bundled omission tier, and one system opt-in tier, with one page count across CLI `convert`, `render`, and `layout`, and Python `to_pdf`, `page_count`, `layout`, and `layout_page`. Diagnostic twins share that public route. Substitution text contains the selected family when a diagnostic channel exists. `inspect` stays deterministic | Route tables, policy helper inside `rdocx` only, CLI branches, Python public calls, `layout_page` | Policy agreement, caller priority, empty slice, one caller font, `layout_page`, diagnostic twins, three commands, derived `Default` false, feature-off explicit system error | Issue 433 case 7 on Linux with DejaVu Sans and the CLI on PATH |

## Test matrix

| Case | Modes or state combined | Expected difference or agreement | Wrong implementation rejected |
|---|---|---|---|
| C1 | `TOC1`, `TOC 1`, `toc 1`, non-canonical id named `toc 1`, body Hyperlink | TOC plain. Body coloured and underlined | A recogniser that misses `TOC1`, or a call into `rdocx::field` |
| C2 | Four indent layers, public, EPUB, ODT paragraph, same `w:ind`, `tbRl` inherited pair | Four at 72.0 pt. ODT run properties unchanged. `tbRl` clears only the absent sibling | Direct-only fix, clearing `CT_RPr`, or leaving `engine.rs:4802` stale |
| C3 | Style-only grid, direct all-`none` | 3 and 3, then 0 and 0 | Copying `none` onto cell borders |
| C4 | No alignment, top, center, bottom | Same split point | Removing `vAlign`, or splitting `cantSplit` |
| C5 | Dark, light, none, explicit black, paragraph shading, nested cell, control paragraph, body wrapper | White on dark, black on light. Nested cell does not inherit. Control paragraph uses the forwarded fill. Body `None` stays black | Paragraph-only fill, painting from the outer cell, or passing `None` at `:1735` or `:1869` |
| C6 | Word symbol path and the six PowerPoint call sites | Word maps Wingdings 2. PowerPoint diagnostics and symbols stay | A shared-function edit, or appending into `current_font_diagnostics` |
| C7 | No caller, empty caller, one caller, omission and system opt-in | Same tier and page count per folder across CLI `convert` and Python `to_pdf`, `page_count`, and `layout` | Fixing only CLI or only Python |
| C7b | CLI `convert`, `render`, and `layout`, flag off and flag on, empty and one caller font | Each command follows its route-table column. `inspect` stays on `layout_deterministic` | Setting `use_system_fonts` on `layout_with_fonts_and_bundled_fallback` and leaving the call unchanged |
| C8 | Caller over embedded and bundled | Caller wins for bundled omission and for the system opt-in | A fallback tier beating the caller |
| C9 | `to_pdf` against `to_pdf_with_diagnostics`, PNG, all pages, and `render_pages`, empty slice and one font, omission and system opt-in | Sibling bytes and font tier agree. Bundled nonempty does not call `layout_with_additional_fonts_and_options` | Leaving a nonempty diagnostic twin on additional fonts, or calling the private helper from Python |
| C10 | `layout` and `layout_page`, Rust and Python, omission and one caller font | Same public route and same tier. Omission returns the shared cache `Arc` | A keyword on `layout` only, or `Arc` built by cloning the inner result |
| C11 | `system-fonts` off, and wasm | An explicit system request errors for an empty font list and for a nonempty font list where that call graph exists. Omission stays on the deterministic cache. Wasm `toPdf` is that same deterministic cache | A silent success for an explicit system request, or describing selected wasm output as the unselected normal engine |
| C12 | Held and fresh layout, save and reopen | Same policy. Substitution facts keep the requested family | A stale cache, or a fact visible to PowerPoint |
| C13 | Unmodelled child beside `w:ind` | `capture_element` bytes preserved | A merge that rewrites unknown XML |
| C14 | Deterministic helpers with the field set both ways | Bundled faces either way | A harness path that reads `use_system_fonts` |

## Rejected alternatives

- Calling `toc_entry_style_ids` from `rdocx-layout`. The function is private and the dependency edge is forbidden.
- Clearing indent only in the layout engine. Public, EPUB, and ODT paragraph reads would keep the stale sibling.
- Changing `odt.rs:1195`. That is run-property merge.
- Copying `none` onto cell borders. Cell `none` is unset and falls back to the table edge.
- Removing `vAlign` itself.
- Treating the donor luminance formula as a pinned LibreOffice file.
- Putting Wingdings 2 into `resolve_symbol_text` or `unknown_symbol_diagnostics`.
- Folding substitution facts into `current_font_diagnostics`. PowerPoint already reads it.
- Counting those additions as two methods.
- Leaving diagnostic methods on `layout_with_additional_fonts_and_options` when the slice is empty.
- Treating an empty slice as a caller-font set.
- Omitting `layout_page`, `render_all_pages`, or `save_pdf` from the keyword or the field.
- A policy enum in `oxml-layout`.
- A new public Rust selector method. The existing `RenderOptions` argument is the consumer for omission inside `rdocx`. Python and CLI nonempty routes select an existing public method in the route table.
- Calling `layout_with_font_policy` from `rdocx-py` or `rdocx-cli`.
- Setting `use_system_fonts` on an ignore-list method and treating that as the tier switch.
- Cloning the inner `WordLayoutResult` to satisfy an owned return.
- A manual `Default`. Selected `RenderOptions` keeps derived `Default`, so `use_system_fonts` is false.
- Python `bundled_fonts` or CLI `--bundled-fonts`.
- Changing caller-only, bundled-fallback, alias, additional-fonts, or deterministic methods so they start honouring `use_system_fonts`.
- Bumping oxml-layout from the unpublished 0.15.0 line to 0.16.0 because `publish = true`.
- Describing the shared public API as unchanged after the three doc-hidden methods.
- Describing selected wasm `toPdf` as bundled normal-engine output. Under Alternative A it is the deterministic cache.
- Refusing every Rust control. Alternative A would then have no zero-font system call.
- Adopting donor hash values, README sizes, or CI results.
- Recording today's system-font output as a deterministic baseline.
- A new source file, module, crate, trait, generic, test binary, or feature flag.

## Test plan

| Timing | Category | Test | Asserts |
|---|---|---|---|
| local scoped | regression | Six `engine.rs` tests | Old-source RED and fixed-source GREEN for A278-1 through A278-6 |
| local scoped | regression | Public, EPUB, and ODT paragraph indent tests, plus the `tbRl` pair at `engine.rs:4802` and `:4805` | The absent sibling clears. ODT run properties do not. Unit rounding does not change |
| local scoped | regression | `dense_form_matches_reviewed_one_page_geometry` update | Border bands move. Donor numbers are not copied |
| local scoped | integration | Existing `rdocx-py/tests/test_core.py` policy test and existing `rdocx-cli` integration module | Same public route for each diagnostic pair, empty slice, and one caller font. CLI `convert`, `render`, and `layout` each discriminate flag off and flag on. `inspect` stays deterministic |
| local scoped | regression | Existing `rdocx` document tests for `Default`, cache identity, and feature-off | Derived `Default` has `use_system_fonts` false. Omission returns the shared deterministic `Arc`. An explicit system request, including a nonempty caller font on the diagnostic and CLI graphs, returns the named error from `cargo test -p rdocx --no-default-features --lib` |
| local scoped | round-trip | `capture_element` beside `w:ind` | Unknown subtree bytes survive |
| final union | differential | Issue 433 cases 1 through 7 | Seven rows, one tier and one page count per folder |
| final union | oracle | LibreOffice `26.2.5.2` `cd7284b4cbbfeb507e630c1aac019f4157393acb` and pdftoppm `26.01.0` from `scripts/docx_ssim_harness.py:39-40`, plus native Word | TOC ink, auto colour, Wingdings square, row split. Native Word version is not pinned |
| final union | schema and font | ECMA-376 controls and the bounded map | Retained XML and child order. Coverage stays unmeasured |
| final union | release riders | Python 3.9 and 3.12, typing and stubs, WASM, archive, hosted CI | Once over the union |

The test gate is differential. Local tests are the implementation gate after a later approved claim. No whole-suite or oracle run is repeated per worker. No executed-test claim is made here. Rendering, native and pinned oracle observations, hash effects, installed Python runtime and typing, wheels, WASM, archives, and hosted CI remain pending.

## HLD impact

Replace the named sentences with the selected Alternative A paragraphs below. Do not append a second policy beside them. The unselected Alternative B paragraphs are not operative.

Files: `docs/hld/08-rendering-spec.md`, `docs/hld/10-bindings-spec.md`, `docs/hld/15-build-and-toolchain.md`. Also add HLD08 sections for TOC ink, first-line indent, table border precedence, row split, automatic text colour, and the Word-only Wingdings 2 map. Add one HLD10 subsection that names `use_system_fonts` and the Python and CLI keywords. No README measurement here. F-X275 owns those measurements.

### `08-rendering-spec.md` cache and Python paragraphs

`Document` keeps a bundled-omission cache and a system-opt-in cache. `layout`, `layout_with_options`, `render_page_to_png`, `render_page_to_svg`, `render_pages`, `render_all_pages`, `to_pdf`, `save_pdf`, and `layout_page` share the bundled cache when `RenderOptions.use_system_fonts` is false. False is derived `Default`. `layout_page` and `layout_page_with_options` keep their route through `layout_with_options`. The same methods share the normal cache when the field is true. Their diagnostic twins use that same choice for an empty caller font slice and return the cached `Arc`. `Arc::clone` preserves sharing. A nonempty caller list is caller, then embedded, then the selected tier, and it is wrapped with `Arc::new` only when the private helper calls an owned method. That `Arc` is not the omission cache. Deterministic methods ignore the field and keep the bundled cache. `layout_with_fonts` stays caller-only. `layout_with_fonts_and_bundled_fallback` and both alias methods stay caller plus bundled. `layout_with_additional_fonts_and_options` stays caller plus the normal engine and ignores the field.

Without selected caller fonts, Python `Document.layout()` and `Document.layout_page()` read the bundled snapshot when `system_fonts` is false. `system_fonts` true reads the normal snapshot, and a build without `system-fonts` returns the documented error. Empty sequences and valid empty directories are omission. Nonempty caller fonts keep caller priority and use the public route table for that same tier. They do not call the private helper and they do not call the normal font methods when the tier is bundled. `page_count`, `to_pdf`, `to_pdf_with_diagnostics`, and every render helper, including `render_all_pages` and its diagnostic twin, use that same keyword and the same public route.

### `10-bindings-spec.md` lines 3191-3216

Native renderers obtain omission output through `Document::layout` and `Document::layout_with_options`. `layout_page` and `layout_page_with_options` keep their route through `layout_with_options` and clone one page out of the returned `Arc`. Accepted omission calls return a shared `Arc<WordLayoutResult>` from the bundled cache, including pages, font bytes, revision view, and the result-local Word source map. `Arc::clone` preserves that sharing. Accepted calls with `use_system_fonts` true return the shared normal-font cache instead. Tracked calls stay uncached and use the engine selected by that field. Nonempty caller-font results that the private helper wraps use `Arc::new` and are not that shared cache. `Document::layout_with_fonts` and `Document::layout_with_fonts_and_options` stay caller-only and uncached. The bundled-fallback methods, including both alias methods, stay caller plus bundled. `layout_with_additional_fonts_and_options` stays caller plus the normal engine and ignores `use_system_fonts`. Deterministic calls remain on the bundled path and ignore the field. These are pre-1.0 native Rust APIs. Python receives `system_fonts`, default false, on the implicit publishers and calls the public route table for a nonempty font list. WASM does not receive the field. `toPdf` follows `Document::to_pdf` into the deterministic cache.

### `10-bindings-spec.md` lines 3909-3920

The constructor and the listed method names remain stable. `toPdf` calls `Document::to_pdf` at `rdocx-wasm/src/lib.rs:85` and returns those bytes directly. Under this alternative that call is the bundled deterministic cache. It does not load host fonts. `rdocx-wasm` disables `rdocx` defaults at `Cargo.toml:27`, while bundled font data stays unconditional. The wasm32 graph therefore excludes host font discovery. No wasm font-policy option is added. A future request for host fonts must fail with the native system-font error. No binding route change.

### `15-build-and-toolchain.md` lines 64-76

Deterministic constructors and the hash harness, golden PNG gate, and SSIM harness use the bundled path. Existing omission constructors and rendering methods also use bundled fonts. `RenderOptions.use_system_fonts` true is the explicit system request. That request loads system fonts when the feature is on and returns the documented error when the feature is off. Deterministic methods ignore the field. Installing or removing system fonts cannot change a bundled omission result.

## Risk routing

Matched rows of `.claude/skills/risk-routing.md` and the extra check each one adds. These riders run once in the final union. They are pending.

- Unit conversion and `Twips`. Read `docs/hld/01-glossary.md` and `CLAUDE.md` truncation note. No constructor rounding change. Declare the hash delta.
- Theme colour. Read `docs/hld/05-drawingml-model.md`. Do not change `rdocx_oxml::theme::apply_tint_shade`. The luminance helper is separate.
- Layout, pagination, line breaking, and shaping. Read `docs/hld/08-rendering-spec.md`. Every baseline uses deterministic font mode.
- Parser and serialiser. Read `docs/hld/04-opc-and-packaging.md` and `docs/hld/06-presentationml-model.md`. `merge_from` can change authored `w:ind`. Preserve child order and raw custody. Add the `capture_element` byte-for-byte round trip.
- Bundled fonts. Read `docs/hld/15-build-and-toolchain.md`. No new asset or licence.
- Public API. Read `docs/hld/10-bindings-spec.md` and `CLAUDE.md` structural rules. The `RenderOptions` field, the Python keywords, the CLI flags, the Word `render_page` visibility change, and exactly three doc-hidden `FontManager` methods. The three methods are the shared 0.15.0 inventory. The field, keywords, flags, and visibility change are Word 0.17.0. Neither line is bumped again. The shared API is changed. PowerPoint behaviour is unchanged. `cargo publish --dry-run` and the size assertion belong to the final gate. No publication here.
- WASM and PyO3. Read `docs/hld/10-bindings-spec.md`. `cargo check --target wasm32-unknown-unknown`. Workspace tests exclude `rdocx-py` and `rpptx-py`.
- Feature default. `cargo test -p oxml-layout --no-default-features` covers the shared bundled database only. The explicit Alternative A system error is `cargo test -p rdocx --no-default-features --lib`. CLI feature-off is `cargo test -p rdocx-cli --no-default-features`. `cargo test -p rdocx-py --no-default-features` does not disable the forced `system-fonts` dependency at `rdocx-py/Cargo.toml:34`.
- External oracle. Pin LibreOffice `26.2.5.2` `cd7284b4cbbfeb507e630c1aac019f4157393acb` and pdftoppm `26.01.0` from `scripts/docx_ssim_harness.py:39-40`. Native Word for Issue 433 is unpinned future evidence.
- Compatibility. Read the design validation rules. Indent units stay. The font default is the answered product choice.
- Cached facts. One cache per tier. Caller-font bundles stay uncached. Substitution facts keep the requested family.
- Shared semantic behaviour. The four issue entry points, plus `layout_page` and each diagnostic twin, agree.
- Contributed issue and PR. Identities pinned. The donor patch is not acceptance.
- Performance. The existing deterministic budget stays 64 MiB and 250 pages per second. No new benchmark.

## Hash harness

The donor delta names 14 keys. Those donor values are not acceptance. Re-record only after the integrated union, in deterministic font mode, with a reviewed expected-output account for each key. The hash effect remains pending.

- `contract:pdf/bytes`, `contract:pdf/pages`
- `feature_showcase:pdf/bytes`, `feature_showcase:pdf/pages`
- `invoice:page1.png`, `invoice:pdf/bytes`, `invoice:pdf/pages`
- `proposal:pdf/bytes`, `proposal:pdf/pages`
- `quote:page1.png`, `quote:pdf/bytes`, `quote:pdf/pages`
- `report:pdf/bytes`, `report:pdf/pages`

`report:page1.png` is unchanged at the donor pin `pr-440.diff:1116`. `dense_form` is a local regression expectation, not a hash key. A278-7 does not move the harness, because deterministic entry points ignore `use_system_fonts`.

## Implementation checklist

- [x] Revalidate the cited callable and consumer contracts against the implementation Base. Current whole-story peers and Root final record that revalidation. Frozen numeric lines remain historical pins, not current line claims.
- [x] Q-S96-07 is answered by the user as Alternative A.
- [x] Implement A278-1 through A278-6 in existing files.
- [x] Implement the private policy helper with an `Arc` return, derived `Default` false, and the three doc-hidden methods.
- [x] Point every Rust diagnostic method at the helper for empty and nonempty slices.
- [x] Point every Python nonempty sibling at the public route table, not at the private helper and not at an ignore-list normal method when the tier is bundled.
- [x] Branch CLI `convert` at `:833` and `:857`, `render` at `:2248`, and `layout` at `:438`. Leave `inspect` on `layout_deterministic`.
- [x] Give Python `layout_page` and every implicit render helper the same keyword and the same public route as `layout`.
- [x] Widen `svg::render_page` to `pub` and re-export it. Do not add a second SVG engine.
- [x] Add local RED and GREEN tests in existing modules, including public, EPUB, ODT, diagnostic twins, and `capture_element`.
- [x] Replace the named HLD paragraphs with the selected Alternative A text.
- [x] Run the scoped gate and microscope after claim.
- [ ] Record the reviewed 14-key hash delta in deterministic font mode.
- [ ] Run the final union once, including F-X273 and F-X274 affected rows.
- [x] Leave README measurements to F-X275.

## Open questions

| ID | Affected F-IDs and A-IDs | Decision and recommendation | Status and answer/source |
|---|---|---|---|
| Q-S96-07 | F-X278, A278-7 only | Bundled omission and explicit system opt-in. Rust `use_system_fonts` defaults false. Python `system_fonts=False`. CLI `--system-fonts` defaults off. Caller fonts stay highest. Empty input is omission. Feature-off explicit system requests return the named error, including nonempty caller fonts on the graphs named above. The exact recorded answer is `Bundled default; system fallback opt-in` | Answered by the user. Record `/private/tmp/rdocx-cross-review/S96_USER_DECISIONS_20261101.json`, SHA-256 `d5caf180fb06aa62d4f55ae9408fbebfc22c42c2d0c3762b268b6d6a50e3e420`. This answer does not approve the contract |

No blocking row remains open. Answering Q-S96-07 does not approve this draft.

## Residuals

- R-A278-1. TOC detection is a style id, a style name, and a `_Toc` anchor. It is not the TOC field. Producer ids outside those forms are outside the rule.
- R-A278-2. Both attributes on one `w:ind` keep hanging precedence. Native Word calibration is pending.
- R-A278-3. Partial edges and conditional style borders have no sample and no count.
- R-A278-4. A short cell is aligned inside the first split part. Word placement is pending.
- R-A278-5. Run shading, highlighting, and a nested cell without its own fill do not supply the background. Word's threshold is undocumented.
- R-A278-6. The Wingdings 2 table is a bounded map, not measured coverage.
- R-A278-7b. `page_count`, `layout`, and `layout_page` discard diagnostics. CLI `layout` does not print them. CLI `convert` and `to_pdf_with_diagnostics` are the channels.
- R-A278-7c. Every diagnostic and non-diagnostic sibling, and CLI `convert`, `render`, and `layout`, has a named public route for a nonempty font. The ID stays so the obligation is auditable.
- R-A278-7d. Desktop omission and wasm `toPdf` both use the deterministic cache. Wasm has no font-policy option and no binding route change.

Raw XML and unmodelled subtrees stay. Child order stays. No new XML element is created.

Refusals that stay in force: an explicit system request on a `system-fonts`-off build returns the named error, including a nonempty caller font on the named diagnostic and CLI graphs. A wasm font-policy option is out of this story. Donor hash values, donor archive sizes, and donor CI are not acceptance. This story does not write a README measurement. A partial-edge count and a guessed Word threshold stay residuals.

## Compatibility for 0.17.0 / 0.15.0

Word 0.17.0 public changes, on the first unpublished Word line: `RenderOptions.use_system_fonts` default false, Python keyword-only `system_fonts: bool = False` on every implicit publisher in the Python table, CLI `--system-fonts` default off on `convert`, `render`, and `layout`, and `svg::render_page` changed from `pub(crate)` to `pub` with `pub use svg::render_page`.

Shared 0.15.0 public additions, on the first unpublished shared line: the three `#[doc(hidden)]` `FontManager` methods in Approach F. No oxml-layout 0.16.0 bump. The shared public API is changed by those three methods. PowerPoint call sites and PowerPoint symbol output stay unchanged. `rpptx` and `rpptx-py` behaviour stays unchanged.

`inspect` gains no flag. Passing `--system-fonts` to `inspect` is an unsupported-flag refusal. `to_pdfa_deterministic` gains no keyword. Ignore-list Rust methods keep today's engines. No publication and no README measurement belong to this story. F-X275 owns the family measurements. This plan does not edit `CHANGELOG.md` or the release scripts.

## Historical approval prerequisites at adoption

At adoption this file stayed `draft` until the prerequisites below were resolved. The operative rolling approval above supersedes this historical readiness description.

1. Semantic closure, any remaining material product question, and the global batch approval are still required. Root final 3 accepted the preparation for adoption and revalidation. It did not approve this canonical design.
2. Three design peers are used and zero remain. Root final follows each phase. A future implementation has a separate cap of three. This adoption is not a counter reset and is not implementation review.
3. Q-S96-07 is answered. The answer does not approve the plan.
4. A clean Base is required before `/claim-feature` and before a shared approved design commit. It is not required for this draft.
5. HLD 14, `BACKLOG.md`, `CURRENT_SPRINT.md`, and `SPRINT_PLAN.md` already list F-X278 as pending and unclaimed, with all 44 A-IDs registered. This adoption does not edit them.
6. Actual rendering, native and pinned oracle observations, the 14-key hash delta, installed Python runtime and typing, wheels, WASM, archives, and hosted CI remain pending.
7. Source F-X272 remains rejected on C1 and C4 and paused for design clarification. It is not a provider. F-X276 remains draft.
8. Actual DeepSeek adoption review and Root final follow before any approval.
9. After a later approval and a clean Base, claim is a separate lease. The worker's first command would be `/start-feature F-X278 --claimed`. This adoption does not start it.

## Evidence pins

| Fact | Value |
|---|---|
| Preparation V6 | `/private/tmp/rdocx-grok-harness/f-x278-preparation/GROK_PROPOSAL_V6.md`, SHA-256 `c1b425a1076d2a7d1e6c74bdd27453498c0fdcf6fa9e7d0821e56bf62fae6f94` |
| Handoff V6 | SHA-256 `0bf0b96d2729bf2469c3b4dc0f60480af9bbdc6804693e71a8200de3cee7953c` |
| Root final 3 | `.claude/reviews/F-X278-codex-preparation-final-3.md`, SHA-256 `875c0ac092d6d2e92b008acc702c57d80433615a68e27055a79532044ac68069` |
| User decision record | `/private/tmp/rdocx-cross-review/S96_USER_DECISIONS_20261101.json`, SHA-256 `d5caf180fb06aa62d4f55ae9408fbebfc22c42c2d0c3762b268b6d6a50e3e420` |
| Frozen source packet | `/private/tmp/rdocx-cross-review/s96-design-text-source-packet`, head `e2a31e58a9b0af20f0ec374e3561c2e599abb0cd`, 919 files, kind `as-of-stopped-source-text-only-not-clean-worktree-or-approval` |
| Issue 433 body | SHA-256 `26bec5a440ed0270568d15e698007ecf78fd74b57b31aa37dd9aabc81386fee0` |
| PR 440 historical head | `94b8c9f15c62de00e355208ed590a7100e8263c1`, base `24af95d3d94c25e95f61a75e7e21f18fa95b6363`, captured `2026-10-31T16:09:30Z` |

Scoped current-worktree reads on 2026-11-01, not a full citation audit:

- `RenderOptions` at `crates/rdocx/src/document.rs:75-80` still has only `revision_view` and `revision_style`. `use_system_fonts` is absent. This does not invalidate the additive field.
- `resolve_run_color` still ends at `Color::BLACK` for `auto` at `crates/rdocx-layout/src/engine.rs:16529`.
- `cell.v_align.is_some()` is still the bail at `crates/rdocx-layout/src/paginator.rs:1844`.
- `layout_deterministic` is still at `crates/rdocx/src/document.rs:45824`.

No accepted preparation clause was invalidated by those reads. Every other line citation remains a frozen 919-packet pin. Implementation-Base revalidation before claim stays required. This adoption did not audit the later 923-file packet and does not treat either packet as a clean Base.


## Local preparation checkpoint, 2026-11-02

All three implementation peers and Codex final2 have completed. Final2 reports0defects0smells0nits after RF278-01 remediation. Implementation checklist records actual local completion. Current callable/consumer revalidation is the recorded source audit after implementation and does not invent a historical preclaim read. The resumed preparation consumes that applicable evidence. No product criterion, public signature, HLD impact or budget changed in this metadata reconciliation.

Plan stays approved and delivery in-progress. The two final checklist items above and all named final union/native/installed/packaging/hosted riders remain pending. Expected14-key changes require actual deterministic integrated observation and an individually reviewed account, not donor values. README measurements remain with F-X275. Publication requires its separate final approval.
