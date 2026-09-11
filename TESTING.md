# Testing Sheety

Sheety compiles spreadsheets (YAML or .xlsx) into standalone Crystal
binaries with an interactive TUI. Three layers of testing cover the
pipeline, each catching failures the previous one can't.

## Running the suite

```bash
shards install       # first time only
crystal spec         # ~340 examples, finishes in seconds
```

## What the specs cover

| Area | Spec files | Notes |
|---|---|---|
| Tokenizer & parser | `formulas_token_spec.cr`, `formulas_parser_spec.cr`, `cell_refs_spec.cr`, `functions_spec.cr`, plus intersection/union/array-constant/named-range specs | Ported from the Python `formulas` library's test suite, so the parser has an upstream oracle. Includes regression tests tied to real past bugs (multi-digit cell refs misclassified as named ranges, XFD1048576, ...). |
| Range normalization | `cell_refs_spec.cr`, `code_generator_spec.cr` | `CellRefs.parse_range` ($ anchors, reversed bounds, whole-column clamping, whole-row rejection) and the generated code it produces; unsupported ranges and unparseable/named-range formulas must surface in `CroupierGenerator#validation_problems` instead of failing silently. |
| Code generation | `code_generator_spec.cr` | Exact/substring assertions on the Crystal emitted for each Excel function, shape deduplication (shared `calc_shape_N` helpers), 2D table arguments (`fetch_cell_range_2d`), arity/type degradation to `#VALUE!`/`#NAME?`, and the range cap. |
| Function registry | `code_generator_spec.cr` (behavioral sections) | Direct calls into `Sheety::Functions.*`: math, text, logical, lookup (VLOOKUP/HLOOKUP approximate *and* exact), conditional aggregation (COUNTIF/SUMIF), dates, `flatten`, plus Excel-fidelity regressions (SQRT of negatives, RIGHT/LEFT edge cases, DATEDIF sign, MIN/MAX of nothing). |
| Shared formulas | `formula_extractor_spec.cr` | `FormulaExtractor.translate_shared`: offset translation of dragged-down formulas ($ anchors, string literals, function names, sheet names, off-sheet #REF!). |
| CLI arg parsing | `cli_spec.cr` | docopt parsing of `<file>`/`--save-to` in any position, usage errors, help/version exits. |
| Generated-program helpers | `croupier_helpers_spec.cr` | Requires `croupier_helpers.cr` directly. This file is **only compiled inside generated binaries** (nothing in `src/` requires it), so without this spec `crystal spec` would never even type-check it — that gap once hid a `Bool#upcase` bug that broke every generated build. |
| Shared utilities | `column_utils_spec.cr`, YAMLParser sections in `code_generator_spec.cr` | Column letter<->number round-trips, YAML scalar typing, `_ui_state` UUID surgery. |
| Roundtrip | `roundtrip_spec.cr` | YAML -> XLSX -> YAML through the real importers and exporter, using temp files. |
| Range cap | sections in `code_generator_spec.cr` | Ranges expanding past `DependencyExtractor::MAX_RANGE_CELLS` (65536) raise `FormulaError`; the generator degrades them to `#VALUE!` tasks and lists them in its validation warning block. |

## What the specs do NOT cover

- The interactive TUI (`tui.cr`), CLI flag handling, the build cache and
  the in-process rebuilder have no automated specs. They are exercised by:
- The CI end-to-end job (`.github/workflows/ci.yml`), which is the only
  automated place where `tui.cr`, `rebuilder.cr` and `croupier_helpers.cr`
  actually get **compiled**: it runs `bin/sheety` on a small sheet, requires
  "Built successfully" in the output (the generated binary's `crystal build`),
  verifies the second run hits the binary cache, and checks `--save-to=.cr`
  conversion.
- Manual runs in a real terminal for TUI behavior.

## Also run locally before declaring done

```bash
shards build                 # build all targets
crystal tool format --check src spec
ameba                        # config in .ameba.yml (excludes generated examples/)
```
