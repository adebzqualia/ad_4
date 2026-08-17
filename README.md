# POPS structural anomaly detector

This project compares each POPS workbook sent to a country with the workbook
received back under the same filename. It produces a self-contained HTML report
for every country and a global dashboard.

Version 0.1 checks one anomaly category:

- **Structural anomalies — HIGH**
  - sheets added or deleted;
  - rows inferred added or deleted;
  - columns inferred added or deleted;
  - sheet type changes;
  - structures that cannot be aligned reliably and therefore require review.

Values, formula edits, formatting edits, hidden/unhidden rows, and changed row
heights or column widths are deliberately not reported as anomalies yet. They
are used only as supporting evidence for row and column alignment.

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
4. builds coordinate-independent fingerprints from labels, normalized formula
   shapes, styles, cell types, row/column properties, merges, tables, filters,
   validations, and drawing anchors;
5. aligns rows while ignoring column positions, then columns while ignoring row
   positions, using stable anchors and an affine-gap sequence alignment;
6. preserves every contiguous insertion/deletion block, including combinations
   with a zero net delta;
7. validates the global row/column map against unique surviving cell anchors.
   Contradictory evidence—often caused by “insert cells down/right” in only part
   of a sheet—is reported as `STRUCTURE_UNRESOLVED`, not as a fictitious whole
   row or column edit.

Every confirmed/inferred structural finding has severity `HIGH`. Confidence is
reported separately as `HIGH`, `MEDIUM`, or `LOW` and reflects positional
evidence, not business impact.

## Important limits

Some histories are mathematically indistinguishable from the final `.xlsx`:

- inserting a completely blank trailing row or column may leave no trace;
- deleting and recreating a row at the same coordinate can look like overwriting
  its cells;
- cut/paste can produce the same final coordinates as insertion/deletion;
- if formulas, values, labels, and structure are all rebuilt, too few anchors
  may survive to prove a global map;
- a local rectangular cell shift is not a whole-row or whole-column edit.

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
stale worksheet dimensions, missing/unexpected files, and report generation.

## Project layout

```text
src/pops_anomaly_detector/
  ooxml.py          secure, read-only workbook structure extraction
  alignment.py      row/column inference and separability validation
  analysis.py       file pairing and result aggregation
  reporting.py      self-contained HTML and JSON output
  cli.py            command-line interface
tests/               synthetic OOXML fixtures and regression tests
```

