# POPS workbook anomaly detector

This project compares each POPS workbook sent to a country with the workbook
received back under the same filename. It produces a self-contained HTML report
for every country and a global dashboard.

Version 0.3 checks four anomaly categories:

- **Structural anomalies — HIGH**
  - sheets added or deleted;
  - rows inferred added or deleted;
  - columns inferred added or deleted;
  - sheet type changes;
  - structures that cannot be aligned reliably and therefore require review.
- **KPI integrity anomalies**
  - missing or unexpected identifiers in the column headed `KPI` on the `KPI`
    sheet — **HIGH**;
  - missing, ambiguous, or non-comparable KPI identifiers — **HIGH**;
  - order-only KPI changes when membership and duplicate counts still match —
    **MEDIUM**.
- **Formula integrity anomalies**
  - an increase in formulas containing an explicit, unquoted `#REF!` token —
    **HIGH**. Stored/cached `#REF!` cells are still counted and shown per sheet,
    but a cache-only change does not create an anomaly because Excel may simply
    have recalculated an unchanged formula;
  - a sent formula removed, replaced by a hardcoded value, or fundamentally
    modified — **MEDIUM**.
- **Value integrity anomalies — MEDIUM**
  - a meaningful prefilled literal from the sent template was changed or
    removed. Blank cells, numeric/text zero (`0`, `0.0`, and equivalent zero
    forms), and a lone `-` are treated as fillable placeholders and are exempt.

Formatting edits, hidden/unhidden rows, and changed row heights or column
widths are deliberately not reported as content anomalies. They remain useful
supporting evidence for row and column alignment.

## Quick start

Python 3.11 or newer is required. The detector has no runtime dependencies.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

Place matching `.xlsx` files here:

```text
data/
  sent/
    France.xlsx
    Germany.xlsx
  received/
    France.xlsx
    Germany.xlsx
```

Run the comparison from the project root:

```powershell
pops-anomaly
```

Or without using the installed console command:

```powershell
python -m pops_anomaly_detector
```

Open `reports/index.html`. Generated files are:

```text
reports/
  index.html                 # global report
  report.json                # machine-readable full result
  countries/
    <country-id>.html        # one report per paired or unpaired file
    <country-id>.json
```

Input workbooks and generated reports are ignored by Git because they may
contain sensitive country data.

## Command-line options

```text
--sent-dir PATH              default: data/sent
--received-dir PATH          default: data/received
--output-dir PATH            default: reports
--recursive                  search below both input directories
--max-active-rows N          default: 100000
--max-active-columns N       default: 16384
--max-cells-per-sheet N      default: 5000000
--max-comparison-cells-per-sheet N
                              default: 500000
--kpi-header-scan-rows N     default: 200
--max-kpi-semantic-cells N   default: 250000
--alignment-band N           default: 240
--always-zero                return 0 even when the report contains ERROR files
```

Files are paired by case-insensitive filename. A duplicate filename in a
recursive input tree is rejected because the intended pair would be ambiguous.
Excel temporary lock files beginning with `~$` are ignored.

Exit codes are suitable for scheduled or CI runs:

- `0`: reports generated and every country is `OK`;
- `1`: reports generated, but at least one country is `ERROR`;
- `2`: invalid input/configuration, report-write failure, or no input files.

Use `--always-zero` if a scheduler should always continue after a completed
report, while still showing the errors in HTML and JSON.

## What “row count” and “column count” mean

Every Excel worksheet always has 1,048,576 possible rows and 16,384 possible
columns. An `.xlsx` file does **not** contain an edit log saying that somebody
clicked “Insert Row.” Consequently, `max_row`, `max_column`, and the cached XML
`<dimension>` value are not reliable anomaly detectors.

The reports use these more precise terms:

- **active row/column span**: the last coordinate supported by stored cells,
  meaningful row/column properties, merged cells, tables, or other bounded
  structural ranges;
- **content rows/columns**: distinct coordinates containing a value or formula;
- **declared dimension**: Excel’s cached metadata, shown for audit only and
  never accepted as proof of an insertion/deletion;
- **added/deleted**: operations inferred by a monotone alignment of stable
  structural evidence;
- **net delta**: additions minus deletions;
- **gross changes**: additions and deletions kept separately. A file with one
  row added and one deleted remains `ERROR` even though its net delta is zero.

## Detection method

The implementation reads the OOXML package directly and never saves or changes
an input workbook. It:

1. resolves workbook parts through package relationships instead of assuming
   worksheet filenames;
2. resolves shared and inline strings and ignores formula cache values;
3. compares semantic style definitions rather than workbook-local style IDs;
4. builds coordinate-independent fingerprints from stable labels, normalized
   formula shapes, styles, cell types, row/column properties, merges, static
   tables, filters, validations, and drawing anchors. Recalculated array/data
   table outputs and refresh-driven PivotTable/query-table cells and ranges are
   not used as structural anchors;
5. aligns rows while ignoring column positions, then columns while ignoring row
   positions, using stable anchors and an affine-gap sequence alignment;
6. preserves every contiguous insertion/deletion block, including combinations
   with a zero net delta;
7. validates the global row/column map against unique surviving cell anchors.
   Contradictory evidence—often caused by “insert cells down/right” in only part
   of a sheet—is reported as `STRUCTURE_UNRESOLVED`, not as a fictitious whole
   row or column edit.
8. on the sheet whose normalized name is exactly `KPI`, locates one literal
   `KPI` header within the configured header scan, reads every nonblank stored
   scalar below it (without stopping at internal blanks), and compares typed,
   normalized identifier sequences with duplicate multiplicity preserved;
9. counts a cell once when it stores the Excel error `#REF!`, its formula
   contains an unquoted `#REF!` token, or both. Text values and quoted formula
   text such as `IFERROR(A1,"#REF!")` are not counted. Only an increase in the
   explicit formula-token component creates the HIGH anomaly; cached errors
   remain audit metrics;
10. compares sent formulas only after both worksheet axes have a validated
    logical map. A1 references are translated into the sent workbook's logical
    coordinates, including cross-sheet references, so Excel's automatic shifts
    after a row or column insertion/deletion do not look like manual edits;
11. reports formula-to-value replacement, formula removal, and changes to
    functions, operators, constants, or safely mapped references. Formula
    result caches are never used as formula identity;
12. compares meaningful sent literals at the same validated logical cells,
    using typed values so numeric `1` equals `1.0`, while deliberately allowing
    countries to complete blank, zero, or `-` placeholders. Stored outputs
    covered by array formulas, data tables, PivotTables, query-backed tables,
    and table calculated columns/totals are excluded because Excel may refresh
    or recalculate them automatically. Those recalculated/generated outputs
    are likewise excluded from row/column inference, while their formula
    anchors remain auditable.

Every confirmed/inferred structural finding has severity `HIGH`. Confidence is
reported separately as `HIGH`, `MEDIUM`, or `LOW` and reflects positional
evidence, not business impact.

Formula removals, hardcoded replacements, fundamental formula edits, and
prefilled-value changes have severity `MEDIUM`. They are aggregated by type and
sheet, with up to 40 representative A1 comparisons in the report. Cells on a
deleted axis or a sheet whose topology cannot be mapped are counted as
skipped/unresolved and never guessed at a physical coordinate.

KPI values are type-aware: numeric `1` and `1.0` match, while text `"1"` is a
different identifier. Literal values have high evidence confidence. A formula
result is used only when the workbook stores a scalar cache and is labelled
medium-confidence; an absent or error cache produces a HIGH unresolved finding
instead of being guessed.

## Important limits

Some histories are mathematically indistinguishable from the final `.xlsx`:

- inserting a completely blank trailing row or column may leave no trace;
- deleting and recreating a row at the same coordinate can look like overwriting
  its cells;
- cut/paste can produce the same final coordinates as insertion/deletion;
- if formulas, values, labels, and structure are all rebuilt, too few anchors
  may survive to prove a global map;
- a local rectangular cell shift is not a whole-row or whole-column edit.
- a read-only OOXML parser does not calculate formulas, follow external links,
  or discover errors that Excel would create only during recalculation;
- unsupported formula syntax, references into an inserted/deleted axis, and
  ambiguous local shifts are skipped rather than classified as manual edits;
- validated shared-formula followers can prove changes to the expression shape
  (for example an operator or constant), but uncertain reference-only changes
  may be omitted to avoid false positives;
- a formula whose `#REF!` text appears only through a shared-formula follower
  may require its stored error cache to be visible at that cell;
- when the KPI sheet contains zero or multiple literal `KPI` header candidates
  in the configured scan region, the comparison is marked unresolved rather
  than choosing a column heuristically.

The detector handles these conservatively. It explains weak evidence, or marks
the sheet unresolved and the country `ERROR`, rather than declaring an
unverifiable workbook `OK`.

For future templates, the strongest improvement would be protected stable UUIDs
for logical rows and columns in hidden cells or defined names. Those identifiers
make otherwise ambiguous replacements deterministic.

## Tests

The test suite builds minimal XLSX packages directly, so Excel and third-party
workbook libraries are not required:

```powershell
python -m unittest discover -s tests -v
```

The cases cover sheet operations, middle and boundary row/column operations,
net-zero combinations, simultaneous row and column edits, content-only changes,
stale worksheet dimensions, exact sheet-name inventories, KPI membership,
duplicates, ordering and unresolved values, `#REF!` lexical edge cases,
formula-cache changes, shared formulas, automatic A1 shifts, formula removal or
hardcoding, meaningful/placeholder value changes, missing/unexpected files,
report filtering, and report generation.

## Project layout

```text
src/pops_anomaly_detector/
  ooxml.py          secure, read-only workbook structure extraction
  alignment.py      row/column inference and separability validation
  formula_logic.py  conservative logical formula normalization
  analysis.py       file pairing and result aggregation
  reporting.py      self-contained HTML and JSON output
  cli.py            command-line interface
tests/               synthetic OOXML fixtures and regression tests
```
