"""Accessible, dependency-free HTML and JSON report generation.

The renderer deliberately keeps every report portable: styles and the small amount
of progressive-enhancement JavaScript are embedded in each HTML document.  All
workbook-derived text is escaped before it reaches HTML.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from collections import OrderedDict
from dataclasses import asdict
from html import escape
from pathlib import Path
from typing import Iterable, Mapping
from uuid import uuid4

from .coordinates import index_to_column
from .models import (
    AxisOperation,
    CountryResult,
    FileEvidence,
    Finding,
    RunResult,
    SheetComparison,
    SheetMetrics,
)

__all__ = ["render_country_report", "render_global_report", "write_reports"]


_SAFE_REPORT_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*\.html$")

_FINDING_LABELS = {
    "SHEET_ADDED": "Sheet added",
    "SHEET_DELETED": "Sheet deleted",
    "SHEET_TYPE_CHANGED": "Sheet type changed",
    "ROWS_INSERTED": "Rows added",
    "ROWS_DELETED": "Rows deleted",
    "COLUMNS_INSERTED": "Columns added",
    "COLUMNS_DELETED": "Columns deleted",
    "STRUCTURE_UNRESOLVED": "Structure unresolved",
    "KPI_HEADER_MISSING": "KPI header missing",
    "KPI_HEADER_AMBIGUOUS": "KPI header ambiguous",
    "KPI_IDENTIFIER_MISSING": "KPI missing",
    "KPI_IDENTIFIER_UNEXPECTED": "Unexpected KPI",
    "KPI_IDENTIFIER_UNRESOLVED": "KPI identifier unresolved",
    "KPI_ORDER_CHANGED": "KPI order changed",
    "REFERENCE_ERRORS_INCREASED": "#REF! errors increased",
    "FORMULA_REPLACED_WITH_VALUE": "Formula replaced with value",
    "FORMULA_REMOVED": "Formula removed",
    "FORMULA_MODIFIED": "Formula modified",
    "PREFILLED_VALUE_CHANGED": "Prefilled value changed",
}

_FINDING_ORDER = {
    "SHEET_ADDED": 10,
    "SHEET_DELETED": 20,
    "SHEET_TYPE_CHANGED": 30,
    "ROWS_INSERTED": 40,
    "ROWS_DELETED": 50,
    "COLUMNS_INSERTED": 60,
    "COLUMNS_DELETED": 70,
    "STRUCTURE_UNRESOLVED": 80,
    "KPI_HEADER_MISSING": 110,
    "KPI_HEADER_AMBIGUOUS": 120,
    "KPI_IDENTIFIER_MISSING": 130,
    "KPI_IDENTIFIER_UNEXPECTED": 140,
    "KPI_IDENTIFIER_UNRESOLVED": 150,
    "KPI_ORDER_CHANGED": 160,
    "REFERENCE_ERRORS_INCREASED": 210,
    "FORMULA_REPLACED_WITH_VALUE": 220,
    "FORMULA_REMOVED": 230,
    "FORMULA_MODIFIED": 240,
    "PREFILLED_VALUE_CHANGED": 310,
}

_CATEGORY_LABELS = {
    "STRUCTURAL": "Structural anomalies",
    "KPI_INTEGRITY": "KPI integrity anomalies",
    "FORMULA_INTEGRITY": "Formula integrity anomalies",
    "VALUE_INTEGRITY": "Value integrity anomalies",
}

_CATEGORY_ORDER = {
    "STRUCTURAL": 10,
    "KPI_INTEGRITY": 20,
    "FORMULA_INTEGRITY": 30,
    "VALUE_INTEGRITY": 40,
}

_SEVERITY_ORDER = {
    "CRITICAL": 0,
    "HIGH": 10,
    "MEDIUM": 20,
    "WARNING": 30,
    "LOW": 40,
    "INFO": 50,
}

_CSS = r"""
:root {
  color-scheme: light;
  --ink: #101828;
  --muted: #475467;
  --subtle: #667085;
  --line: #d0d5dd;
  --line-soft: #eaecf0;
  --surface: #ffffff;
  --canvas: #f6f8fb;
  --accent: #175cd3;
  --accent-soft: #eff8ff;
  --ok: #067647;
  --ok-soft: #ecfdf3;
  --error: #b42318;
  --error-soft: #fee4e2;
  --warn: #b54708;
  --warn-soft: #fffaeb;
  --shadow: 0 1px 2px rgba(16, 24, 40, .06);
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; max-width: 100%; overflow-x: clip; }
body {
  margin: 0;
  max-width: 100%;
  overflow-x: clip;
  color: var(--ink);
  background: var(--canvas);
  font: 16px/1.5 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
a { color: var(--accent); text-underline-offset: .16em; }
a:hover { text-decoration-thickness: 2px; }
a:focus-visible, button:focus-visible, input:focus-visible, select:focus-visible,
summary:focus-visible, [tabindex]:focus-visible {
  outline: 3px solid rgba(23, 92, 211, .35);
  outline-offset: 2px;
}
.skip-link {
  position: fixed; z-index: 100; top: .5rem; left: .5rem;
  padding: .65rem .9rem; color: #fff; background: var(--ink);
  transform: translateY(-160%);
}
.skip-link:focus { transform: translateY(0); }
.topbar { color: #fff; background: #102a56; border-bottom: 4px solid #2e90fa; }
.topbar__inner, .page, .footer__inner {
  width: min(1440px, calc(100% - 2rem)); margin-inline: auto;
}
.topbar__inner { padding: 1.05rem 0; }
.brand { margin: 0; font-size: .82rem; font-weight: 750; letter-spacing: .08em; text-transform: uppercase; }
.breadcrumbs, .report-nav { display: flex; flex-wrap: wrap; gap: .45rem .8rem; align-items: center; }
.breadcrumbs { margin-top: .4rem; color: #d1e9ff; font-size: .9rem; }
.breadcrumbs a { color: #fff; }
.page { padding: 2rem 0 3.5rem; }
.hero { display: flex; gap: 1.25rem; justify-content: space-between; align-items: flex-start; margin-bottom: 1.5rem; }
.hero h1 { margin: 0 0 .35rem; font-size: clamp(1.75rem, 4vw, 2.55rem); line-height: 1.15; letter-spacing: -.025em; }
.eyebrow { margin: 0 0 .45rem; color: var(--accent); font-size: .78rem; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; }
.lede { max-width: 78ch; margin: 0; color: var(--muted); }
.hero__status { display: flex; flex-wrap: wrap; gap: .5rem; justify-content: flex-end; }
.section { margin-top: 2rem; scroll-margin-top: 1rem; }
.section__heading { display: flex; flex-wrap: wrap; gap: .75rem; justify-content: space-between; align-items: baseline; margin-bottom: .8rem; }
.section h2 { margin: 0; font-size: 1.3rem; }
.section h3 { margin: 0; font-size: 1.05rem; }
.section-intro { max-width: 90ch; margin: .2rem 0 1rem; color: var(--muted); }
.grid { display: grid; gap: 1rem; }
.kpi-grid { grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); }
.two-col { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.card {
  min-width: 0; padding: 1rem; background: var(--surface); border: 1px solid var(--line-soft);
  border-radius: .65rem; box-shadow: var(--shadow);
}
.kpi { border-top: 3px solid #98a2b3; }
.kpi--ok { border-top-color: var(--ok); }
.kpi--error { border-top-color: var(--error); }
.kpi--warn { border-top-color: var(--warn); }
.kpi--accent { border-top-color: var(--accent); }
.kpi__label { color: var(--muted); font-size: .82rem; font-weight: 700; text-transform: uppercase; letter-spacing: .035em; }
.kpi__value { margin-top: .22rem; font-size: 1.65rem; font-weight: 780; font-variant-numeric: tabular-nums; line-height: 1.2; }
.kpi__help { margin-top: .32rem; color: var(--subtle); font-size: .82rem; }
.badge {
  display: inline-flex; gap: .34rem; align-items: center; width: fit-content;
  padding: .2rem .55rem; border: 1px solid currentColor; border-radius: 999px;
  font-size: .75rem; font-weight: 800; letter-spacing: .025em; line-height: 1.4; white-space: nowrap;
}
.badge--ok { color: var(--ok); background: var(--ok-soft); }
.badge--error, .badge--high { color: var(--error); background: var(--error-soft); }
.badge--warn, .badge--medium { color: var(--warn); background: var(--warn-soft); }
.badge--info { color: #175cd3; background: var(--accent-soft); }
.badge--neutral { color: #344054; background: #f2f4f7; }
.notice { padding: .9rem 1rem; border: 1px solid var(--line); border-left-width: 5px; border-radius: .45rem; background: var(--surface); }
.notice + .notice { margin-top: .65rem; }
.notice--error { border-left-color: var(--error); background: #fff8f7; }
.notice--warn { border-left-color: var(--warn); background: #fffdf5; }
.notice--ok { border-left-color: var(--ok); background: #f6fef9; }
.notice--info { border-left-color: var(--accent); background: #f7fbff; }
.notice__title { margin: 0 0 .25rem; font-weight: 780; }
.notice p, .notice ul { margin-top: .3rem; margin-bottom: 0; }
.toolbar {
  display: grid; grid-template-columns: minmax(220px, 2fr) minmax(180px, 1fr) auto;
  gap: .75rem; align-items: end; margin: 1rem 0;
}
.anomaly-toolbar { grid-template-columns: repeat(2, minmax(190px, 1fr)) auto auto; }
.field label { display: block; margin-bottom: .3rem; color: var(--muted); font-size: .84rem; font-weight: 700; }
.field input, .field select {
  width: 100%; min-height: 2.65rem; padding: .55rem .7rem; color: var(--ink); background: #fff;
  border: 1px solid #98a2b3; border-radius: .4rem; font: inherit;
}
.filter-count { min-height: 2.65rem; display: flex; align-items: center; color: var(--muted); font-variant-numeric: tabular-nums; }
.table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: .55rem; background: #fff; box-shadow: var(--shadow); }
table { width: 100%; border-collapse: collapse; }
caption { padding: .75rem 1rem; color: var(--muted); font-size: .86rem; text-align: left; }
th, td { padding: .72rem .8rem; border-top: 1px solid var(--line-soft); text-align: left; vertical-align: top; }
thead th { color: #344054; background: #f9fafb; border-top: 0; font-size: .78rem; letter-spacing: .025em; text-transform: uppercase; }
tbody th { font-weight: 700; }
tbody tr:hover { background: #fcfcfd; }
.global-table { min-width: 1080px; }
.sheet-table { min-width: 1440px; }
.cell-title { font-weight: 700; }
.cell-meta, .muted { color: var(--subtle); font-size: .84rem; }
.cell-meta { margin-top: .18rem; }
.number { font-variant-numeric: tabular-nums; white-space: nowrap; }
.delta { display: block; font-variant-numeric: tabular-nums; white-space: nowrap; }
.delta--changed { color: var(--error); font-weight: 700; }
.button-link {
  display: inline-flex; min-height: 2.35rem; align-items: center; justify-content: center;
  padding: .45rem .75rem; border: 1px solid #84adff; border-radius: .4rem;
  color: #1849a9; background: #f5f8ff; font-weight: 700; text-decoration: none;
}
button.button-link { font: inherit; cursor: pointer; }
.report-nav { margin: 0 0 1rem; }
.report-nav__spacer { flex: 1; }
.progress-card label { display: block; margin-bottom: .45rem; font-weight: 700; }
progress { width: 100%; height: .75rem; accent-color: var(--accent); }
.progress-card__text { margin: .45rem 0 0; color: var(--muted); }
.anomaly-type {
  margin-top: .75rem; overflow: hidden; background: var(--surface);
  border: 1px solid var(--line); border-radius: .65rem; box-shadow: var(--shadow);
}
.anomaly-type > summary, .anomaly-sheet > summary {
  width: 100%; padding: .78rem 1rem; color: var(--ink);
}
.anomaly-type > summary { background: #f9fafb; }
.anomaly-type[open] > summary { border-bottom: 1px solid var(--line-soft); }
.anomaly-type__summary-content, .anomaly-sheet__summary-content {
  display: inline-flex; flex-wrap: wrap; gap: .45rem; align-items: center;
  margin-left: .3rem;
}
.anomaly-type__body { padding: 0 1rem 1rem; }
.anomaly-sheet { margin: .8rem 0 0; border: 1px solid var(--line-soft); border-radius: .5rem; }
.anomaly-sheet > summary { background: #fcfcfd; }
.anomaly-sheet > .finding-list { padding: 0 .8rem .8rem; }
.finding-list { display: grid; gap: .7rem; margin-top: .7rem; }
.finding { border-left: 5px solid var(--error); }
.finding[data-severity="MEDIUM"], .finding[data-severity="WARNING"] { border-left-color: var(--warn); }
.finding[data-severity="LOW"], .finding[data-severity="INFO"] { border-left-color: var(--accent); }
.finding__top { display: flex; flex-wrap: wrap; gap: .45rem; align-items: center; }
.finding__title { margin-right: auto; font-size: 1rem; }
.finding__message { margin: .7rem 0; }
.mini-dl, .audit-dl { display: grid; grid-template-columns: max-content minmax(0, 1fr); gap: .25rem .8rem; margin: .5rem 0 0; }
.mini-dl dt, .audit-dl dt { color: var(--muted); font-weight: 650; }
.mini-dl dd, .audit-dl dd { min-width: 0; margin: 0; overflow-wrap: anywhere; }
.evidence-list, .notes-list { margin: .45rem 0 0; padding-left: 1.2rem; }
.finding[data-category="FORMULA_INTEGRITY"] .evidence-list,
.finding[data-category="VALUE_INTEGRITY"] .evidence-list {
  display: grid; gap: .35rem; padding-left: 0; list-style: none;
  font: .82rem/1.5 ui-monospace, SFMono-Regular, Consolas, monospace;
}
.finding[data-category="FORMULA_INTEGRITY"] .evidence-list li,
.finding[data-category="VALUE_INTEGRITY"] .evidence-list li {
  padding: .45rem .55rem; overflow-wrap: anywhere; white-space: pre-wrap;
  border: 1px solid var(--line-soft); border-radius: .35rem; background: #f9fafb;
}
details { margin-top: .55rem; }
summary { width: fit-content; color: #344054; cursor: pointer; font-weight: 700; }
.sheet-evidence-row > td { padding: .35rem .8rem .8rem; background: #fcfcfd; }
.sheet-evidence-row:hover { background: #fcfcfd; }
.sheet-evidence { margin: 0; }
.sheet-evidence .mini-dl { grid-template-columns: minmax(6rem, max-content) minmax(0, 1fr); }
.sheet-evidence .mini-dl dd { overflow-wrap: break-word; word-break: normal; }
.reference-errors { font-variant-numeric: tabular-nums; white-space: nowrap; }
.reference-errors--present { color: var(--error); font-weight: 700; }
.cell-integrity { font-variant-numeric: tabular-nums; white-space: nowrap; }
.cell-integrity .cell-meta { white-space: normal; }
.cell-integrity__item { display: block; }
.cell-integrity__item--changed { color: var(--warn); font-weight: 700; }
.sheet-inventory { padding: 0; }
.sheet-inventory > summary { width: 100%; padding: 1rem; }
.sheet-inventory[open] > summary { border-bottom: 1px solid var(--line-soft); }
.sheet-inventory__summary { display: inline-flex; flex-wrap: wrap; gap: .5rem; align-items: center; margin-left: .3rem; }
.sheet-inventory__body { padding: 1rem; }
.inventory-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 1rem; }
.inventory-panel { min-width: 0; padding: .8rem; border: 1px solid var(--line-soft); border-radius: .5rem; }
.inventory-panel h3 { margin: 0 0 .5rem; }
.inventory-list { margin: 0; padding-left: 2rem; }
.inventory-list li { padding: .18rem 0; overflow-wrap: break-word; }
.inventory-list li::marker { color: var(--subtle); font-variant-numeric: tabular-nums; }
.inventory-name { margin-right: .4rem; }
.kpi-evidence { margin-top: .7rem; padding-top: .65rem; border-top: 1px solid var(--line-soft); }
.kpi-evidence h4 { margin: 0 0 .35rem; font-size: .95rem; }
.kpi-entry-list { max-height: 18rem; overflow: auto; margin: .35rem 0 0; padding-left: 1.25rem; }
.analysis-notes { margin-top: 0; }
.analysis-notes > summary { width: 100%; }
.operation-list { margin: .35rem 0 0; padding-left: 1.05rem; }
.operation-list li + li { margin-top: .25rem; }
.source-card h3 { margin-bottom: .65rem; }
.hash { font: .78rem/1.45 ui-monospace, SFMono-Regular, Consolas, monospace; overflow-wrap: anywhere; }
.method-list { max-width: 95ch; margin: .5rem 0 0; padding-left: 1.25rem; }
.method-list li + li { margin-top: .45rem; }
.footer { border-top: 1px solid var(--line); color: var(--muted); background: #fff; }
.footer__inner { padding: 1.1rem 0; font-size: .84rem; }
.print-only { display: none; }
.sr-only {
  position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
  overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;
}
[hidden] { display: none !important; }
@media (max-width: 760px) {
  .topbar__inner, .page, .footer__inner { width: min(100% - 1rem, 1440px); }
  .page { padding-top: 1.25rem; }
  .hero { display: block; }
  .hero__status { justify-content: flex-start; margin-top: .8rem; }
  .two-col, .toolbar, .inventory-grid { grid-template-columns: 1fr; }
  .filter-count { min-height: auto; }
}
@media (prefers-reduced-motion: reduce) {
  html { scroll-behavior: auto; }
}
@media print {
  @page { size: A4 landscape; margin: 11mm; }
  :root { --canvas: #fff; --shadow: none; }
  body { font-size: 9pt; background: #fff; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .screen-only, .skip-link, .toolbar { display: none !important; }
  .print-only { display: block; }
  .page, .topbar__inner, .footer__inner { width: 100%; }
  .page { padding: 1rem 0; }
  .topbar { color: #000; background: #fff; border-bottom: 2px solid #000; }
  .breadcrumbs { color: #333; }
  .breadcrumbs a { color: #000; }
  .card, .table-wrap { box-shadow: none; }
  .table-wrap { overflow: visible; }
  .global-table, .sheet-table { min-width: 0; }
  thead { display: table-header-group; }
  tr, .card, .finding, .notice, .source-card { break-inside: avoid; }
  .section { break-inside: auto; }
  .badge { color: #000 !important; background: #fff !important; border: 1.5px solid #000; }
  details:not([open]) > *:not(summary) { display: block !important; }
  [data-finding][hidden], [data-sheet-group][hidden], [data-type-group][hidden],
  .anomaly-category[hidden] { display: block !important; }
  summary { list-style: none; }
  summary::-webkit-details-marker { display: none; }
  a { color: #000; text-decoration: underline; }
}
"""

_GLOBAL_FILTER_JS = r"""
(() => {
  const search = document.getElementById('country-search');
  const filter = document.getElementById('result-filter');
  const count = document.getElementById('filter-count');
  const rows = Array.from(document.querySelectorAll('#country-results tbody tr'));
  if (!search || !filter || !count) return;
  const update = () => {
    const query = search.value.trim().toLocaleLowerCase();
    const selected = filter.value;
    let visible = 0;
    rows.forEach((row) => {
      const matchesText = !query || row.dataset.search.toLocaleLowerCase().includes(query);
      const matchesStatus = selected === 'ALL' || row.dataset.status === selected || row.dataset.state === selected;
      row.hidden = !(matchesText && matchesStatus);
      if (!row.hidden) visible += 1;
    });
    count.textContent = `${visible} of ${rows.length} results shown`;
  };
  search.addEventListener('input', update);
  filter.addEventListener('change', update);
  update();
})();
"""

_ANOMALY_FILTER_JS = r"""
(() => {
  const root = document.getElementById('anomaly-results');
  const typeFilter = document.getElementById('anomaly-type-filter');
  const severityFilter = document.getElementById('anomaly-severity-filter');
  const reset = document.getElementById('anomaly-filter-reset');
  const count = document.getElementById('anomaly-filter-count');
  if (!root || !typeFilter || !severityFilter || !reset || !count) return;

  const cards = Array.from(root.querySelectorAll('[data-finding]'));
  const sheetGroups = Array.from(root.querySelectorAll('[data-sheet-group]'));
  const typeGroups = Array.from(root.querySelectorAll('[data-type-group]'));
  const categories = Array.from(root.querySelectorAll('.anomaly-category[data-category]'));

  const update = () => {
    const selectedType = typeFilter.value;
    const selectedSeverity = severityFilter.value;
    const filtering = selectedType !== 'ALL' || selectedSeverity !== 'ALL';
    let visible = 0;

    cards.forEach((card) => {
      const typeMatches = selectedType === 'ALL' || card.dataset.type === selectedType;
      const severityMatches = selectedSeverity === 'ALL' || card.dataset.severity === selectedSeverity;
      card.hidden = !(typeMatches && severityMatches);
      if (!card.hidden) visible += 1;
    });

    sheetGroups.forEach((group) => {
      const groupCount = group.querySelectorAll('[data-finding]:not([hidden])').length;
      group.hidden = groupCount === 0;
      const output = group.querySelector('[data-sheet-visible-count]');
      if (output) output.textContent = String(groupCount);
      if (filtering && groupCount) group.open = true;
    });

    typeGroups.forEach((group) => {
      const groupCount = group.querySelectorAll('[data-finding]:not([hidden])').length;
      group.hidden = groupCount === 0;
      const output = group.querySelector('[data-type-visible-count]');
      if (output) output.textContent = String(groupCount);
      if (filtering && groupCount) group.open = true;
    });

    categories.forEach((category) => {
      const categoryCount = category.querySelectorAll('[data-finding]:not([hidden])').length;
      category.hidden = filtering && categoryCount === 0;
    });
    count.textContent = `${visible} of ${cards.length} anomalies shown`;
  };

  const openAncestors = (target) => {
    for (let node = target ? target.parentElement : null; node; node = node.parentElement) {
      if (node instanceof HTMLDetailsElement) {
        node.hidden = false;
        node.open = true;
      }
      if (node.classList && node.classList.contains('anomaly-category')) node.hidden = false;
    }
  };

  const revealHash = () => {
    if (!window.location.hash) return;
    const target = document.getElementById(decodeURIComponent(window.location.hash.slice(1)));
    if (!target || !target.matches('[data-finding]')) return;
    typeFilter.value = 'ALL';
    severityFilter.value = 'ALL';
    update();
    openAncestors(target);
  };

  typeFilter.addEventListener('change', update);
  severityFilter.addEventListener('change', update);
  reset.addEventListener('click', () => {
    typeFilter.value = 'ALL';
    severityFilter.value = 'ALL';
    update();
    typeFilter.focus();
  });
  document.addEventListener('click', (event) => {
    const link = event.target.closest ? event.target.closest('a[href^="#finding-"]') : null;
    if (!link) return;
    typeFilter.value = 'ALL';
    severityFilter.value = 'ALL';
    update();
    openAncestors(document.getElementById(link.hash.slice(1)));
  });
  window.addEventListener('hashchange', revealHash);
  update();
  revealHash();
})();
"""

_PRINT_DETAILS_JS = r"""
(() => {
  let closed = [];
  window.addEventListener('beforeprint', () => {
    closed = Array.from(document.querySelectorAll('details:not([open])'));
    closed.forEach((item) => { item.open = true; });
  });
  window.addEventListener('afterprint', () => {
    closed.forEach((item) => { item.open = false; });
    closed = [];
  });
})();
"""


def _e(value: object | None) -> str:
    """Escape a possibly workbook-derived value for HTML text or attributes."""

    return escape("" if value is None else str(value), quote=True)


def _number(value: int) -> str:
    return f"{value:,}"


def _signed(value: int) -> str:
    if value > 0:
        return f"+{value:,}"
    if value < 0:
        return f"−{abs(value):,}"
    return "0"


def _human_bytes(value: int) -> str:
    size = float(value)
    units = ("bytes", "KB", "MB", "GB", "TB")
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "bytes":
                return f"{int(size):,} {unit}"
            return f"{size:,.1f} {unit}"
        size /= 1024
    return f"{value:,} bytes"


def _safe_filename(raw_name: str, fallback: str) -> str:
    if Path(raw_name).name == raw_name and _SAFE_REPORT_NAME.fullmatch(raw_name):
        return raw_name
    normalized = re.sub(r"[^A-Za-z0-9._-]+", "-", fallback).strip(".-_") or "country"
    digest = hashlib.sha256(raw_name.encode("utf-8", errors="replace")).hexdigest()[:8]
    return f"{normalized[:80]}-{digest}.html"


def _filename_map(countries: Iterable[CountryResult]) -> dict[int, str]:
    result: dict[int, str] = {}
    used: set[str] = set()
    for index, country in enumerate(countries):
        name = _safe_filename(country.report_filename, country.country_id or country.display_name)
        stem = Path(name).stem
        suffix = Path(name).suffix
        candidate = name
        discriminator = 2
        while candidate.casefold() in used:
            candidate = f"{stem}-{discriminator}{suffix}"
            discriminator += 1
        used.add(candidate.casefold())
        result[index] = candidate
    return result


def _atomic_write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        temporary.write_text(content, encoding="utf-8", newline="\n")
        os.replace(temporary, path)
    finally:
        if temporary.exists():
            temporary.unlink()


def _badge(label: object, tone: str = "neutral", symbol: str | None = None) -> str:
    allowed_tones = {"ok", "error", "high", "medium", "warn", "info", "neutral"}
    resolved_tone = tone if tone in allowed_tones else "neutral"
    symbol_html = f'<span aria-hidden="true">{_e(symbol)}</span>' if symbol else ""
    return f'<span class="badge badge--{resolved_tone}">{symbol_html}{_e(label)}</span>'


def _status_badge(status: str) -> str:
    normalized = status.upper()
    if normalized in {"OK", "UNCHANGED"}:
        return _badge(normalized, "ok", "✓")
    if normalized in {"ERROR", "MODIFIED", "ADDED", "DELETED"}:
        return _badge(normalized, "error", "!")
    if normalized in {"MISSING_RECEIVED", "UNEXPECTED_RECEIVED", "READ_ERROR"}:
        return _badge(normalized.replace("_", " "), "warn", "!")
    if normalized in {"WARNING", "WARN"}:
        return _badge("WARNING", "warn", "!")
    if normalized == "PAIRED":
        return _badge("PAIRED", "info")
    return _badge(normalized.replace("_", " "))


def _severity_badge(severity: str | None) -> str:
    if not severity:
        return _badge("No anomaly severity", "neutral")
    normalized = severity.upper()
    if normalized in {"CRITICAL", "HIGH"}:
        return _badge(normalized, "high", "!!")
    if normalized in {"MEDIUM", "WARNING"}:
        return _badge(normalized, "medium", "!")
    if normalized in {"LOW", "INFO"}:
        return _badge(normalized, "info")
    return _badge(normalized, "neutral")


def _category_label(category: str) -> str:
    normalized = category.upper()
    return _CATEGORY_LABELS.get(normalized, normalized.replace("_", " ").title())


def _severity_sort_key(severity: str) -> tuple[int, str]:
    normalized = severity.upper()
    return (_SEVERITY_ORDER.get(normalized, 999), normalized)


def _maximum_severity(findings: Iterable[Finding]) -> str | None:
    severities = {finding.severity.upper() for finding in findings if finding.severity}
    return min(severities, key=_severity_sort_key) if severities else None


def _category_results(country: CountryResult) -> list[dict[str, object]]:
    grouped: dict[str, list[Finding]] = {category: [] for category in _CATEGORY_ORDER}
    for finding in country.findings:
        grouped.setdefault(finding.category.upper(), []).append(finding)
    results: list[dict[str, object]] = []
    for category in sorted(
        grouped,
        key=lambda value: (_CATEGORY_ORDER.get(value, 999), value),
    ):
        findings = grouped[category]
        maximum = _maximum_severity(findings)
        if not findings and category == "STRUCTURAL" and country.comparison_state != "PAIRED":
            status = "NOT_RUN"
        elif maximum in {"CRITICAL", "HIGH"}:
            status = "ERROR"
        elif maximum in {"MEDIUM", "WARNING"}:
            status = "WARNING"
        elif findings:
            status = "INFO"
        else:
            status = "OK"
        results.append(
            {
                "code": category,
                "label": _category_label(category),
                "severity": maximum,
                "status": status,
                "finding_count": len(findings),
            }
        )
    return results


def _kpi(label: str, value: object, help_text: str, tone: str = "") -> str:
    tone_class = f" kpi--{tone}" if tone in {"ok", "error", "warn", "accent"} else ""
    return (
        f'<article class="card kpi{tone_class}">'
        f'<div class="kpi__label">{_e(label)}</div>'
        f'<div class="kpi__value">{_e(value)}</div>'
        f'<div class="kpi__help">{_e(help_text)}</div>'
        "</article>"
    )


def _notice(title: str, messages: Iterable[str], tone: str) -> str:
    materialized = list(messages)
    if not materialized:
        return ""
    items = "".join(f"<li>{_e(message)}</li>" for message in materialized)
    return (
        f'<aside class="notice notice--{tone}">'
        f'<p class="notice__title">{_e(title)}</p><ul>{items}</ul></aside>'
    )


def _collapsible_notice(title: str, messages: Iterable[str], tone: str) -> str:
    materialized = list(messages)
    if not materialized:
        return ""
    items = "".join(f"<li>{_e(message)}</li>" for message in materialized)
    return (
        f'<details class="analysis-notes notice notice--{_e(tone)}">'
        f'<summary>{_e(title)} ({len(materialized):,})</summary>'
        f'<ul class="notes-list">{items}</ul></details>'
    )


def _document(title: str, body: str, script: str = "") -> str:
    script_html = f"<script>{script}</script>" if script else ""
    return (
        "<!doctype html>\n"
        '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{_e(title)}</title>\n"
        f"<style>{_CSS}</style>\n</head>\n<body>\n"
        '<a class="skip-link" href="#main-content">Skip to main content</a>\n'
        f"{body}\n{script_html}\n</body>\n</html>\n"
    )


def _country_payload(country: CountryResult, report_filename: str) -> dict[str, object]:
    payload = asdict(country)
    payload["report_filename"] = report_filename
    payload["report_href"] = f"countries/{report_filename}"
    payload["category_results"] = _category_results(country)
    metrics = payload.get("metrics")
    if isinstance(metrics, dict):
        metrics["row_net_delta"] = getattr(country.metrics, "row_net_delta", 0)
        metrics["column_net_delta"] = getattr(country.metrics, "column_net_delta", 0)
        metrics["sheet_net_delta"] = getattr(country.metrics, "sheet_net_delta", 0)
    return payload


def _run_payload(run: RunResult, filenames: Mapping[int, str]) -> dict[str, object]:
    payload = run.to_dict()
    payload["countries"] = [
        _country_payload(country, filenames[index])
        for index, country in enumerate(run.countries)
    ]
    return payload


def _finding_label(code: str) -> str:
    return _FINDING_LABELS.get(code, code.replace("_", " ").title())


def _unit_label(code: str, count: int) -> str:
    if code.startswith("ROW"):
        unit = "row"
    elif code.startswith("COLUMN"):
        unit = "column"
    elif code.startswith("SHEET"):
        unit = "sheet"
    elif code.startswith("KPI"):
        unit = "KPI"
    elif "REFERENCE_ERROR" in code or "REF_ERROR" in code:
        unit = "cell"
    elif code.startswith("FORMULA") or code.startswith("PREFILLED_VALUE"):
        unit = "cell"
    else:
        unit = "case"
    return f"{count:,} {unit if count == 1 else unit + 's'}"


def _finding_rollups(run: RunResult) -> list[dict[str, object]]:
    buckets: dict[tuple[str, str, str], dict[str, object]] = {}
    for country in run.countries:
        for finding in country.findings:
            category = finding.category.upper()
            severity = finding.severity.upper()
            key = (category, finding.code, severity)
            bucket = buckets.setdefault(
                key,
                {
                    "category": category,
                    "code": finding.code,
                    "severity": severity,
                    "findings": 0,
                    "units": 0,
                    "countries": set(),
                    "sheets": set(),
                },
            )
            bucket["findings"] = int(bucket["findings"]) + 1
            bucket["units"] = int(bucket["units"]) + max(0, finding.unit_count)
            countries = bucket["countries"]
            sheets = bucket["sheets"]
            if isinstance(countries, set):
                countries.add(country.country_id)
            sheet_name = finding.sent_sheet_name or finding.received_sheet_name
            if sheet_name and isinstance(sheets, set):
                sheets.add((country.country_id, sheet_name))
    return sorted(
        buckets.values(),
        key=lambda item: (
            _CATEGORY_ORDER.get(str(item["category"]), 999),
            _FINDING_ORDER.get(str(item["code"]), 999),
            _severity_sort_key(str(item["severity"])),
            str(item["code"]),
        ),
    )


def _global_rollup_html(run: RunResult) -> str:
    rollups = _finding_rollups(run)
    if not rollups:
        return (
            '<div class="notice notice--ok"><p class="notice__title">No anomaly findings</p>'
            "<p>No anomalies were detected in the comparable workbooks.</p></div>"
        )
    rows: list[str] = []
    for item in rollups:
        code = str(item["code"])
        category = str(item["category"])
        severity = str(item["severity"])
        countries = item["countries"]
        sheets = item["sheets"]
        country_count = len(countries) if isinstance(countries, set) else 0
        sheet_count = len(sheets) if isinstance(sheets, set) else 0
        findings = int(item["findings"])
        units = int(item["units"])
        rows.append(
            "<tr>"
            f'<td>{_e(_category_label(category))}</td>'
            f'<th scope="row"><span class="cell-title">{_e(_finding_label(code))}</span>'
            f'<div class="cell-meta">{_e(code)}</div></th>'
            f'<td class="number">{findings:,}</td>'
            f'<td class="number">{_e(_unit_label(code, units))}</td>'
            f'<td class="number">{country_count:,}</td>'
            f'<td class="number">{sheet_count:,}</td>'
            f"<td>{_severity_badge(severity)}</td>"
            "</tr>"
        )
    return (
        '<div class="table-wrap" tabindex="0" role="region" aria-label="Anomaly roll-up">'
        '<table><caption>Findings count operations or unresolved cases; affected units count rows, columns, sheets, KPI identifiers, or cells.</caption>'
        "<thead><tr><th scope=\"col\">Category</th><th scope=\"col\">Type</th><th scope=\"col\">Findings</th>"
        "<th scope=\"col\">Affected units</th><th scope=\"col\">Countries</th>"
        "<th scope=\"col\">Sheets</th><th scope=\"col\">Severity</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


def _file_names(country: CountryResult) -> str:
    sent = country.sent_file.name if country.sent_file else "No sent file"
    received = country.received_file.name if country.received_file else "No received file"
    return f"{country.display_name} {sent} {received}"


def _operation_summary(added: int, deleted: int) -> str:
    return f"+{added:,} added / −{deleted:,} deleted (net {_signed(added - deleted)})"


def _global_country_rows(run: RunResult, filenames: Mapping[int, str]) -> str:
    rows: list[str] = []
    for index, country in enumerate(run.countries):
        metrics = country.metrics
        high_count = sum(finding.severity.upper() == "HIGH" for finding in country.findings)
        medium_count = sum(finding.severity.upper() == "MEDIUM" for finding in country.findings)
        comparable = country.comparison_state == "PAIRED"
        sent_name = country.sent_file.name if country.sent_file else "Not present"
        received_name = country.received_file.name if country.received_file else "Not present"
        if comparable:
            sheets = f"{metrics.sent_sheet_count:,} → {metrics.received_sheet_count:,}"
            sheet_meta = _operation_summary(metrics.sheets_added, metrics.sheets_deleted)
            sheet_meta += (
                "; names match exactly"
                if getattr(metrics, "sheet_names_match", True)
                else "; names differ"
            )
            row_change = _operation_summary(metrics.rows_added, metrics.rows_deleted)
            column_change = _operation_summary(metrics.columns_added, metrics.columns_deleted)
        else:
            sheets = "Not compared"
            sheet_meta = row_change = column_change = "Unavailable"
        search_text = _file_names(country)
        href = f"countries/{filenames[index]}"
        rows.append(
            f'<tr data-status="{_e(country.overall_status.upper())}" '
            f'data-state="{_e(country.comparison_state.upper())}" data-search="{_e(search_text)}">'
            f'<th scope="row"><span class="cell-title">{_e(country.display_name)}</span>'
            f'<div class="cell-meta">Sent: {_e(sent_name)}</div>'
            f'<div class="cell-meta">Received: {_e(received_name)}</div></th>'
            f"<td>{_status_badge(country.overall_status)}<div class=\"cell-meta\">"
            f"{_status_badge(country.comparison_state)}</div></td>"
            f'<td class="number">{high_count:,} HIGH'
            f'<div class="cell-meta">{medium_count:,} MEDIUM</div></td>'
            f'<td><span class="number">{_e(sheets)}</span><div class="cell-meta">{_e(sheet_meta)}</div></td>'
            f'<td><span class="number">{_e(row_change)}</span></td>'
            f'<td><span class="number">{_e(column_change)}</span></td>'
            f'<td class="number">{metrics.affected_sheet_count:,}</td>'
            f'<td><a class="button-link" href="{_e(href)}" aria-label="Open report for {_e(country.display_name)}">Open report</a></td>'
            "</tr>"
        )
    if not rows:
        return '<tr><td colspan="8">No workbook files were discovered.</td></tr>'
    return "".join(rows)


def _global_audit_html(run: RunResult) -> str:
    scope = ", ".join(run.scope) if run.scope else "Not specified"
    return (
        '<div class="card"><dl class="audit-dl">'
        f"<dt>Run ID</dt><dd><code>{_e(run.run_id)}</code></dd>"
        f'<dt>Generated</dt><dd><time datetime="{_e(run.generated_at_utc)}">{_e(run.generated_at_utc)}</time></dd>'
        f"<dt>Comparator</dt><dd>{_e(run.comparator_version)}</dd>"
        f"<dt>Schema</dt><dd>{_e(run.schema_version)}</dd>"
        f"<dt>Scope</dt><dd>{_e(scope)}</dd>"
        f"<dt>Sent directory</dt><dd><code>{_e(run.sent_directory)}</code></dd>"
        f"<dt>Received directory</dt><dd><code>{_e(run.received_directory)}</code></dd>"
        "</dl></div>"
    )


def _render_global_report(run: RunResult, filenames: Mapping[int, str]) -> str:
    summary = run.summary
    coverage_max = max(summary.sent_files, 1)
    quality_max = max(summary.matched_pairs, 1)
    scope_labels = {
        "SHEETS": "sheets",
        "ROWS": "rows",
        "COLUMNS": "columns",
        "KPI_IDENTIFIERS": "KPI identifiers",
        "REFERENCE_ERRORS": "#REF! reference errors",
        "FORMULAS": "aligned formulas",
        "PREFILLED_VALUES": "aligned prefilled values",
    }
    scope_text = ", ".join(scope_labels.get(item, item.replace("_", " ").lower()) for item in run.scope)
    scope_text = scope_text or "workbook evidence"
    formula_changes = sum(
        int(getattr(country.metrics, "formula_changed_count", 0))
        for country in run.countries
    )
    formula_unresolved = sum(
        int(getattr(country.metrics, "formula_unresolved_count", 0))
        for country in run.countries
    )
    value_changes = sum(
        int(getattr(country.metrics, "value_changed_count", 0))
        for country in run.countries
    )
    value_unresolved = sum(
        int(getattr(country.metrics, "value_unresolved_count", 0))
        for country in run.countries
    )
    run_severity = (
        "HIGH"
        if summary.high_findings
        else "MEDIUM"
        if getattr(summary, "medium_findings", 0)
        else None
    )
    notices = _notice("Run notes", run.warnings, "warn")
    body = f"""
<header class="topbar">
  <div class="topbar__inner">
    <p class="brand">POPS anomaly detection</p>
    <nav class="breadcrumbs" aria-label="Breadcrumb"><span aria-current="page">Global report</span></nav>
  </div>
</header>
<main id="main-content" class="page">
  <div class="hero">
    <div>
      <p class="eyebrow">Workbook integrity</p>
      <h1>POPS global report</h1>
      <p class="lede">Comparison of {summary.sent_files:,} sent and {summary.received_files:,} received workbooks. This run evaluates {_e(scope_text)} where stored workbook evidence is available.</p>
    </div>
    <div class="hero__status" aria-label="Run status">
      {_status_badge('ERROR' if summary.error else 'OK')}
      {_severity_badge(run_severity)}
    </div>
  </div>
  {notices}
  <section class="section" aria-labelledby="overview-heading">
    <div class="section__heading"><h2 id="overview-heading">Run overview</h2><a href="report.json">Download report JSON</a></div>
    <div class="grid kpi-grid">
      {_kpi('Sent files', _number(summary.sent_files), 'Expected country workbooks', 'accent')}
      {_kpi('Received files', _number(summary.received_files), 'Returned country workbooks', 'accent')}
      {_kpi('Matched pairs', _number(summary.matched_pairs), 'Same filename on both sides', 'accent')}
      {_kpi('OK', _number(summary.ok), f'Of {summary.matched_pairs:,} matched comparisons', 'ok')}
      {_kpi('ERROR', _number(summary.error), 'Workbooks requiring action', 'error' if summary.error else '')}
      {_kpi('HIGH findings', _number(summary.high_findings), f'Across {summary.affected_countries:,} countries', 'error' if summary.high_findings else '')}
      {_kpi('MEDIUM findings', _number(getattr(summary, 'medium_findings', 0)), 'Review recommended', 'warn' if getattr(summary, 'medium_findings', 0) else '')}
      {_kpi('Formula changes', _number(formula_changes), f'{formula_unresolved:,} unresolved or skipped', 'warn' if formula_changes or formula_unresolved else '')}
      {_kpi('Prefilled value changes', _number(value_changes), f'{value_unresolved:,} unresolved or skipped', 'warn' if value_changes or value_unresolved else '')}
    </div>
  </section>
  <section class="section" aria-labelledby="coverage-heading">
    <div class="section__heading"><h2 id="coverage-heading">Coverage and comparison quality</h2></div>
    <div class="grid two-col">
      <article class="card progress-card">
        <label for="coverage-progress">Matched sent workbooks</label>
        <progress id="coverage-progress" max="{coverage_max}" value="{min(summary.matched_pairs, coverage_max)}">{summary.matched_pairs} of {summary.sent_files}</progress>
        <p class="progress-card__text"><strong>{summary.matched_pairs:,} of {summary.sent_files:,}</strong> sent workbooks matched; {summary.missing_received:,} missing received and {summary.unexpected_received:,} unexpected received.</p>
      </article>
      <article class="card progress-card">
        <label for="quality-progress">Matched workbooks with OK result</label>
        <progress id="quality-progress" max="{quality_max}" value="{min(summary.ok, quality_max)}">{summary.ok} of {summary.matched_pairs}</progress>
        <p class="progress-card__text"><strong>{summary.ok:,} of {summary.matched_pairs:,}</strong> matched comparisons are OK; {summary.comparison_failed:,} comparisons require attention because reading or structural mapping failed.</p>
      </article>
    </div>
  </section>
  <section class="section" aria-labelledby="rollup-heading">
    <div class="section__heading"><h2 id="rollup-heading">Anomalies by category and type</h2></div>
    <p class="section-intro">Structural additions and deletions are counted independently. KPI, formula, and prefilled-value findings are reported in separate integrity categories with their own criticality.</p>
    {_global_rollup_html(run)}
  </section>
  <section class="section" aria-labelledby="countries-heading">
    <div class="section__heading"><h2 id="countries-heading">Country results</h2></div>
    <div class="toolbar screen-only" role="search" aria-label="Filter country results">
      <div class="field"><label for="country-search">Search country or filename</label><input id="country-search" type="search" autocomplete="off" placeholder="Start typing…"></div>
      <div class="field"><label for="result-filter">Result or comparison state</label><select id="result-filter"><option value="ALL">All results</option><option value="OK">OK</option><option value="ERROR">ERROR</option><option value="MISSING_RECEIVED">Missing received</option><option value="UNEXPECTED_RECEIVED">Unexpected received</option><option value="READ_ERROR">Read error</option></select></div>
      <div id="filter-count" class="filter-count" role="status" aria-live="polite">{len(run.countries):,} results shown</div>
    </div>
    <noscript><p class="notice notice--info">Search and filters require JavaScript. All results remain visible below.</p></noscript>
    <div class="table-wrap" tabindex="0" role="region" aria-label="Country comparison results">
      <table id="country-results" class="global-table">
        <caption>One row per matched, missing, unexpected, or unreadable country workbook.</caption>
        <thead><tr><th scope="col">Country / files</th><th scope="col">Result</th><th scope="col">Findings</th><th scope="col">Sheets sent → received</th><th scope="col">Row operations</th><th scope="col">Column operations</th><th scope="col">Affected sheets</th><th scope="col"><span class="sr-only">Report link</span></th></tr></thead>
        <tbody>{_global_country_rows(run, filenames)}</tbody>
      </table>
    </div>
  </section>
  <section class="section" aria-labelledby="method-heading">
    <div class="section__heading"><h2 id="method-heading">Methodology and limitations</h2></div>
    <div class="card">
      <ul class="method-list">
        <li>Sheet identity is matched by exact Excel sheet name. A rename is reported conservatively as one deleted sheet and one added sheet.</li>
        <li>Active worksheet extents are derived from stored cells and structural evidence such as row or column properties, ranges, tables, and drawing anchors. Excel's declared dimension is retained as diagnostic evidence rather than trusted on its own.</li>
        <li>Inserted and deleted rows and columns are inferred by aligning structural signatures. Recalculated array/data-table outputs and refresh-driven PivotTable/query-table cells and ranges are excluded from those signatures. Ambiguous mappings are surfaced as a HIGH manual-review finding instead of being silently treated as unchanged.</li>
        <li>Operation counts, not net deltas, determine structural status. Equal additions and deletions may leave the same final dimensions.</li>
        <li>The KPI sheet is checked for an identifiable KPI column and differences in normalized KPI identifiers and order.</li>
        <li>#REF! counts include formula text and cached error cells; per-sheet evidence shows both components, which can overlap. Only an increase in explicit formula tokens creates a HIGH finding because an unchanged formula's cache may be recalculated by Excel.</li>
        <li>Formulas are compared only after logical row and column alignment. Stored cached results are ignored; formulas inside removed axes or sheets with unresolved mappings are skipped rather than compared at the wrong coordinates.</li>
        <li>Meaningful prefilled sent values are compared at logically aligned cells. Blank, zero, and lone-hyphen placeholders, plus stored array-formula, data-table, PivotTable, query-table, and calculated-column/total outputs, are excluded. Cells inside removed axes or unresolved mappings are skipped and counted separately.</li>
        <li>Formatting intent, charts, and broader business-rule validation remain outside this report's scope.</li>
      </ul>
    </div>
  </section>
  <section class="section" aria-labelledby="audit-heading">
    <div class="section__heading"><h2 id="audit-heading">Run evidence</h2></div>
    {_global_audit_html(run)}
  </section>
  <p class="print-only">Run {_e(run.run_id)} · generated {_e(run.generated_at_utc)}</p>
</main>
<footer class="footer"><div class="footer__inner">POPS anomaly detector · Run <code>{_e(run.run_id)}</code> · Generated <time datetime="{_e(run.generated_at_utc)}">{_e(run.generated_at_utc)}</time></div></footer>
"""
    return _document("POPS global anomaly report", body, _GLOBAL_FILTER_JS + _PRINT_DETAILS_JS)


def render_global_report(run: RunResult) -> str:
    """Render the global HTML report without writing it to disk."""

    return _render_global_report(run, _filename_map(run.countries))


def _file_card(label: str, evidence: FileEvidence | None) -> str:
    if evidence is None:
        return (
            '<article class="card source-card">'
            f"<h3>{_e(label)}</h3><p class=\"muted\">File not available.</p></article>"
        )
    return f"""
<article class="card source-card">
  <h3>{_e(label)}</h3>
  <dl class="audit-dl">
    <dt>Filename</dt><dd>{_e(evidence.name)}</dd>
    <dt>Relative path</dt><dd><code>{_e(evidence.relative_path)}</code></dd>
    <dt>Size</dt><dd>{_e(_human_bytes(evidence.size_bytes))} ({evidence.size_bytes:,} bytes)</dd>
    <dt>Modified</dt><dd><time datetime="{_e(evidence.modified_at_utc)}">{_e(evidence.modified_at_utc)}</time></dd>
    <dt>SHA-256</dt><dd><code class="hash">{_e(evidence.sha256)}</code></dd>
  </dl>
</article>
"""


def _axis_position(start: int, end: int, axis: str) -> str:
    normalized = axis.upper()
    if normalized == "COLUMN":
        try:
            first = index_to_column(start)
            last = index_to_column(end)
        except ValueError:
            first = str(start)
            last = str(end)
    else:
        first = str(start)
        last = str(end)
    return first if start == end else f"{first}–{last}"


def _finding_location(finding: Finding) -> str:
    if finding.start is None:
        return "Not applicable"
    end = finding.end if finding.end is not None else finding.start
    if finding.code.upper().startswith("COLUMN"):
        axis = "COLUMN"
    elif finding.code.upper().startswith("ROW"):
        axis = "ROW"
    else:
        axis = "POSITION"
    positions = _axis_position(finding.start, end, axis)
    coordinate = finding.coordinate_space.lower() if finding.coordinate_space else "unspecified"
    return f"{positions} ({coordinate} coordinates)"


def _finding_card(finding: Finding, anchor: str) -> str:
    evidence = ""
    if finding.evidence:
        items = "".join(f"<li>{_e(item)}</li>" for item in finding.evidence)
        evidence = (
            f'<details><summary>Evidence ({len(finding.evidence):,})</summary>'
            f'<ul class="evidence-list">{items}</ul></details>'
        )
    confidence_tone = "info" if finding.confidence.upper() == "HIGH" else "warn"
    sheet_name = finding.sent_sheet_name or finding.received_sheet_name or "Workbook"
    return f"""
<article id="{_e(anchor)}" class="card finding" data-finding
  data-category="{_e(finding.category.upper())}" data-type="{_e(finding.code)}"
  data-severity="{_e(finding.severity.upper())}" data-sheet="{_e(sheet_name)}">
  <div class="finding__top">
    <h4 class="finding__title">{_e(_finding_label(finding.code))}</h4>
    {_severity_badge(finding.severity)}
    {_badge(f'{finding.confidence} confidence', confidence_tone)}
  </div>
  <p class="finding__message">{_e(finding.message)}</p>
  <dl class="mini-dl">
    <dt>Finding ID</dt><dd><code>{_e(finding.id)}</code></dd>
    <dt>Code</dt><dd><code>{_e(finding.code)}</code></dd>
    <dt>Scope</dt><dd>{_e(finding.scope)}</dd>
    <dt>Affected units</dt><dd>{finding.unit_count:,}</dd>
    <dt>Location</dt><dd>{_e(_finding_location(finding))}</dd>
  </dl>
  {evidence}
</article>
"""


def _finding_groups_html(
    findings: Iterable[Finding],
    anchors: Mapping[int, str],
    empty_message: str,
) -> str:
    materialized = list(findings)
    categories: dict[str, list[Finding]] = {"STRUCTURAL": []}
    for finding in materialized:
        categories.setdefault(finding.category.upper(), []).append(finding)
    output: list[str] = []
    for category, category_findings in sorted(
        categories.items(),
        key=lambda item: (_CATEGORY_ORDER.get(item[0], 999), item[0]),
    ):
        category_id = re.sub(r"[^a-z0-9]+", "-", category.casefold()).strip("-") or "other"
        maximum = _maximum_severity(category_findings)
        if category == "STRUCTURAL":
            intro = (
                "Added and deleted operations are independent, including when their net "
                "effect on worksheet dimensions is zero."
            )
        elif category == "KPI_INTEGRITY":
            intro = (
                "The KPI identifier column is compared between the sent and received KPI "
                "sheets, including missing, unexpected, ambiguous, and reordered identifiers."
            )
        elif category == "FORMULA_INTEGRITY":
            intro = (
                "Formula integrity covers #REF! errors and MEDIUM changes where an aligned "
                "sent formula was removed, replaced by a value, or logically modified. "
                "Cached formula results are ignored."
            )
        elif category == "VALUE_INTEGRITY":
            intro = (
                "Prefilled sent values are compared with their logically aligned received "
                "cells. Blank, zero, and lone-hyphen placeholders are fillable and excluded. "
                "Confirmed changes are MEDIUM; cells that cannot be mapped safely are skipped."
            )
        else:
            intro = "Findings in this category are grouped by anomaly type and sheet."

        if not category_findings:
            category_body = (
                '<div class="notice notice--ok"><p class="notice__title">No structural findings</p>'
                f"<p>{_e(empty_message)}</p></div>"
            )
        else:
            by_type: dict[str, list[Finding]] = {}
            for finding in category_findings:
                by_type.setdefault(finding.code, []).append(finding)
            type_html: list[str] = []
            for code, type_findings in sorted(
                by_type.items(),
                key=lambda item: (_FINDING_ORDER.get(item[0], 999), item[0]),
            ):
                sheets: OrderedDict[str, list[Finding]] = OrderedDict()
                for finding in type_findings:
                    sheet_name = (
                        finding.sent_sheet_name or finding.received_sheet_name or "Workbook"
                    )
                    sheets.setdefault(sheet_name, []).append(finding)
                sheet_html: list[str] = []
                for sheet_name, sheet_findings in sheets.items():
                    sheet_label = "Workbook" if sheet_name == "Workbook" else f"Sheet: {sheet_name}"
                    cards = "".join(
                        _finding_card(item, anchors[id(item)]) for item in sheet_findings
                    )
                    sheet_html.append(
                        f'<details class="anomaly-sheet" data-sheet-group data-sheet="{_e(sheet_name)}">'
                        '<summary><span class="anomaly-sheet__summary-content">'
                        f'<span>{_e(sheet_label)}</span>'
                        f'<span class="muted">(<span data-sheet-visible-count>{len(sheet_findings):,}</span>)</span>'
                        f'{_severity_badge(_maximum_severity(sheet_findings))}'
                        '</span></summary>'
                        f'<div class="finding-list">{cards}</div></details>'
                    )
                severity_badges = "".join(
                    _severity_badge(severity)
                    for severity in sorted(
                        {finding.severity.upper() for finding in type_findings},
                        key=_severity_sort_key,
                    )
                )
                type_html.append(
                    f'<details class="anomaly-type" data-type-group data-type="{_e(code)}" open>'
                    '<summary><span class="anomaly-type__summary-content">'
                    f'<span>{_e(_finding_label(code))}</span>'
                    f'<code>{_e(code)}</code>'
                    f'<span class="muted">(<span data-type-visible-count>{len(type_findings):,}</span>)</span>'
                    f'{severity_badges}</span></summary>'
                    f'<div class="anomaly-type__body">{"".join(sheet_html)}</div></details>'
                )
            category_body = "".join(type_html)

        heading_badge = _severity_badge(maximum) if maximum else _status_badge("OK")
        output.append(
            f'<section id="category-{_e(category_id)}" class="section anomaly-category" '
            f'data-category="{_e(category)}" aria-labelledby="category-{_e(category_id)}-heading">'
            f'<div class="section__heading"><h2 id="category-{_e(category_id)}-heading">'
            f'{_e(_category_label(category))}</h2>{heading_badge}</div>'
            f'<p class="section-intro">{_e(intro)}</p>{category_body}</section>'
        )
    return "".join(output)


def _anomaly_filters_html(findings: Iterable[Finding]) -> str:
    materialized = list(findings)
    if not materialized:
        return ""
    codes = sorted(
        {finding.code for finding in materialized},
        key=lambda code: (_FINDING_ORDER.get(code, 999), code),
    )
    severities = sorted(
        {finding.severity.upper() for finding in materialized},
        key=_severity_sort_key,
    )
    type_options = "".join(
        f'<option value="{_e(code)}">{_e(_finding_label(code))}</option>' for code in codes
    )
    severity_options = "".join(
        f'<option value="{_e(severity)}">{_e(severity)}</option>'
        for severity in severities
    )
    return f"""
<div id="anomaly-filters" class="toolbar anomaly-toolbar screen-only" role="search" aria-label="Filter anomalies">
  <div class="field"><label for="anomaly-type-filter">Anomaly type</label><select id="anomaly-type-filter"><option value="ALL">All types</option>{type_options}</select></div>
  <div class="field"><label for="anomaly-severity-filter">Criticality</label><select id="anomaly-severity-filter"><option value="ALL">All criticalities</option>{severity_options}</select></div>
  <button id="anomaly-filter-reset" class="button-link" type="button">Reset filters</button>
  <div id="anomaly-filter-count" class="filter-count" role="status" aria-live="polite">{len(materialized):,} of {len(materialized):,} anomalies shown</div>
</div>
<noscript><p class="notice notice--info">Anomaly filters require JavaScript. All anomalies remain visible below.</p></noscript>
"""


def _metric_html(metrics: SheetMetrics | None) -> str:
    if metrics is None:
        return '<span class="muted">Not present</span>'
    active_ref = metrics.active_ref or "Empty"
    content_ref = metrics.content_ref or "No content"
    return (
        f'<span class="cell-title number">{metrics.active_rows:,} rows × {metrics.active_columns:,} columns</span>'
        f'<div class="cell-meta">Active: {_e(active_ref)}</div>'
        f'<div class="cell-meta">Content: {_e(content_ref)}</div>'
    )


def _extent_delta(comparison: SheetComparison) -> str:
    if comparison.sent_metrics is None or comparison.received_metrics is None:
        return '<span class="muted">Not comparable</span>'
    row_delta = comparison.received_metrics.active_rows - comparison.sent_metrics.active_rows
    column_delta = (
        comparison.received_metrics.active_columns - comparison.sent_metrics.active_columns
    )
    row_class = " delta--changed" if row_delta else ""
    column_class = " delta--changed" if column_delta else ""
    return (
        f'<span class="delta{row_class}">Rows {_e(_signed(row_delta))}</span>'
        f'<span class="delta{column_class}">Columns {_e(_signed(column_delta))}</span>'
    )


def _axis_operation_text(operation: AxisOperation, axis: str) -> str:
    verb = "Added" if operation.operation.upper() == "ADDED" else "Deleted"
    end = operation.end
    position = _axis_position(operation.start, end, axis)
    unit = "row" if axis.upper() == "ROW" else "column"
    unit_label = unit if operation.count == 1 else f"{unit}s"
    return (
        f"{verb} {operation.count:,} {unit_label} at {position}; "
        f"{operation.coordinate_space.lower()} coordinates; {operation.confidence.lower()} confidence"
    )


def _axis_operations_html(
    operations: Iterable[AxisOperation],
    axis: str,
    comparison_status: str,
) -> str:
    materialized = list(operations)
    if not materialized:
        text = "None detected" if comparison_status.upper() == "UNCHANGED" else "None resolved"
        return f'<span class="muted">{_e(text)}</span>'
    added = sum(op.count for op in materialized if op.operation.upper() == "ADDED")
    deleted = sum(op.count for op in materialized if op.operation.upper() == "DELETED")
    items = "".join(
        f"<li>{_e(_axis_operation_text(operation, axis))}</li>"
        for operation in materialized
    )
    return (
        f'<span class="cell-title number">{_e(_operation_summary(added, deleted))}</span>'
        f'<ul class="operation-list">{items}</ul>'
    )


def _metrics_evidence(label: str, metrics: SheetMetrics | None) -> str:
    if metrics is None:
        return f"<dt>{_e(label)}</dt><dd>Not present</dd>"
    declared = metrics.declared_dimension or "Not declared"
    ref_total = int(getattr(metrics, "ref_error_count", getattr(metrics, "ref_error_cells", 0)))
    cached = int(
        getattr(metrics, "cached_ref_error_count", getattr(metrics, "cached_ref_error_cells", 0))
    )
    formula = int(
        getattr(metrics, "formula_ref_error_count", getattr(metrics, "formula_ref_error_cells", 0))
    )
    return (
        f"<dt>{_e(label)}</dt><dd>Declared {_e(declared)}; "
        f"{metrics.populated_cells:,} populated cells; {metrics.formula_cells:,} formulas; "
        f"{metrics.styled_blank_cells:,} styled blanks; {metrics.merged_ranges:,} merged ranges; "
        f"{metrics.table_ranges:,} tables; {ref_total:,} unique #REF! cells "
        f"({cached:,} cached-error cells, {formula:,} formula-text cells)</dd>"
    )


def _reference_errors_html(comparison: SheetComparison) -> str:
    sent_metrics = comparison.sent_metrics
    received_metrics = comparison.received_metrics
    sent = (
        int(getattr(sent_metrics, "ref_error_count", getattr(sent_metrics, "ref_error_cells", 0)))
        if sent_metrics is not None
        else None
    )
    received = (
        int(
            getattr(
                received_metrics,
                "ref_error_count",
                getattr(received_metrics, "ref_error_cells", 0),
            )
        )
        if received_metrics is not None
        else None
    )
    sent_text = "—" if sent is None else f"{sent:,}"
    received_text = "—" if received is None else f"{received:,}"
    present_class = " reference-errors--present" if received else ""
    if sent is None or received is None:
        delta = '<span class="cell-meta">Delta unavailable</span>'
    else:
        difference = received - sent
        delta_class = " delta--changed" if difference else ""
        delta = f'<span class="delta{delta_class}">Delta {_e(_signed(difference))}</span>'
    return (
        f'<span class="reference-errors{present_class}">Sent {sent_text} → '
        f"Received {received_text}</span>{delta}"
    )


def _cell_integrity_html(comparison: SheetComparison) -> str:
    formula_changes = int(getattr(comparison, "formula_changed_count", 0))
    value_changes = int(getattr(comparison, "value_changed_count", 0))
    formula_unresolved = int(getattr(comparison, "formula_unresolved_count", 0))
    value_unresolved = int(getattr(comparison, "value_unresolved_count", 0))
    formula_class = " cell-integrity__item--changed" if formula_changes else ""
    value_class = " cell-integrity__item--changed" if value_changes else ""
    unresolved = formula_unresolved + value_unresolved
    unresolved_html = ""
    if unresolved:
        unresolved_html = (
            f'<span class="cell-meta">Skipped/unresolved: {unresolved:,} '
            f"({formula_unresolved:,} formula / {value_unresolved:,} value)</span>"
        )
    return (
        f'<span class="cell-integrity__item{formula_class}">Formula changes: '
        f"{formula_changes:,}</span>"
        f'<span class="cell-integrity__item{value_class}">Value changes: '
        f"{value_changes:,}</span>{unresolved_html}"
    )


def _cell_integrity_evidence_html(comparison: SheetComparison) -> str:
    formula_changes = int(getattr(comparison, "formula_changed_count", 0))
    value_changes = int(getattr(comparison, "value_changed_count", 0))
    formula_unresolved = int(getattr(comparison, "formula_unresolved_count", 0))
    value_unresolved = int(getattr(comparison, "value_unresolved_count", 0))
    if not any((formula_changes, value_changes, formula_unresolved, value_unresolved)):
        return ""
    return (
        '<div class="kpi-evidence"><h4>Aligned cell-integrity comparison</h4>'
        '<dl class="mini-dl">'
        f"<dt>Formula changes</dt><dd>{formula_changes:,}</dd>"
        f"<dt>Value changes</dt><dd>{value_changes:,}</dd>"
        f"<dt>Formula skipped/unresolved</dt><dd>{formula_unresolved:,}</dd>"
        f"<dt>Value skipped/unresolved</dt><dd>{value_unresolved:,}</dd>"
        "</dl><p class=\"muted\">A1 locations and sent → received formulas or values "
        "are listed in the linked finding evidence.</p></div>"
    )


def _kpi_snapshot_html(label: str, snapshot: object | None) -> str:
    if snapshot is None:
        return f'<div class="inventory-panel"><h4>{_e(label)}</h4><p class="muted">Not available.</p></div>'
    entries = list(getattr(snapshot, "entries", []) or [])
    header = getattr(snapshot, "header_coordinate", None) or "Not resolved"
    candidates = list(getattr(snapshot, "header_candidates", []) or [])
    duplicates = dict(getattr(snapshot, "duplicate_keys", {}) or {})
    notes = list(getattr(snapshot, "notes", []) or [])
    candidate_text = ", ".join(str(item) for item in candidates) if candidates else "None"
    duplicate_count = len(duplicates)
    entry_items = "".join(
        "<li>"
        f'<code>{_e(getattr(entry, "coordinate", ""))}</code> '
        f'{_e(getattr(entry, "display_value", ""))} '
        f'<span class="muted">({_e(getattr(entry, "value_kind", "unknown"))}; '
        f'{_e(getattr(entry, "confidence", "unspecified"))} confidence)</span></li>'
        for entry in entries
    )
    entries_html = (
        f'<details><summary>{_e(label)} KPI identifiers ({len(entries):,})</summary>'
        f'<ol class="kpi-entry-list">{entry_items}</ol></details>'
        if entries
        else ""
    )
    notes_html = (
        '<ul class="notes-list">'
        + "".join(f"<li>{_e(note)}</li>" for note in notes)
        + "</ul>"
        if notes
        else ""
    )
    return (
        f'<div class="inventory-panel"><h4>{_e(label)}</h4><dl class="mini-dl">'
        f'<dt>Status</dt><dd>{_e(getattr(snapshot, "status", "UNKNOWN"))}</dd>'
        f'<dt>KPI header</dt><dd>{_e(header)}</dd>'
        f'<dt>Header candidates</dt><dd>{_e(candidate_text)}</dd>'
        f'<dt>Identifiers</dt><dd>{len(entries):,}</dd>'
        f'<dt>Duplicate keys</dt><dd>{duplicate_count:,}</dd>'
        f"</dl>{entries_html}{notes_html}</div>"
    )


def _kpi_comparison_html(comparison: SheetComparison) -> str:
    kpi = getattr(comparison, "kpi_comparison", None)
    if kpi is None:
        return ""
    sent = getattr(kpi, "sent", None)
    received = getattr(kpi, "received", None)
    return (
        '<div class="kpi-evidence"><h4>KPI identifier comparison</h4>'
        '<dl class="mini-dl">'
        f'<dt>Status</dt><dd>{_e(getattr(kpi, "status", "UNKNOWN"))}</dd>'
        f'<dt>Missing</dt><dd>{int(getattr(kpi, "missing_count", 0)):,}</dd>'
        f'<dt>Unexpected</dt><dd>{int(getattr(kpi, "unexpected_count", 0)):,}</dd>'
        f'<dt>Reordered</dt><dd>{int(getattr(kpi, "reordered_count", 0)):,}</dd>'
        "</dl><div class=\"inventory-grid\">"
        f"{_kpi_snapshot_html('Sent', sent)}{_kpi_snapshot_html('Received', received)}"
        "</div></div>"
    )


def _sheet_evidence_html(comparison: SheetComparison) -> str:
    notes = ""
    if comparison.alignment_notes:
        notes = "<ul class=\"notes-list\">" + "".join(
            f"<li>{_e(note)}</li>" for note in comparison.alignment_notes
        ) + "</ul>"
    open_attr = " open" if comparison.status.upper() != "UNCHANGED" else ""
    return (
        f'<details class="sheet-evidence"{open_attr}><summary>Sheet evidence</summary>'
        '<dl class="mini-dl">'
        f"{_metrics_evidence('Sent', comparison.sent_metrics)}"
        f"{_metrics_evidence('Received', comparison.received_metrics)}"
        f"</dl>{_kpi_comparison_html(comparison)}"
        f"{_cell_integrity_evidence_html(comparison)}{notes}</details>"
    )


def _sheet_name_html(comparison: SheetComparison) -> str:
    sent = comparison.sent_name or "Not present"
    received = comparison.received_name or "Not present"
    sent_meta = (
        f"position {comparison.sent_index}; {comparison.sent_type}"
        if comparison.sent_index is not None
        else "not present"
    )
    received_meta = (
        f"position {comparison.received_index}; {comparison.received_type}"
        if comparison.received_index is not None
        else "not present"
    )
    return (
        f'<span class="cell-title">Sent: {_e(sent)}</span><div class="cell-meta">{_e(sent_meta)}</div>'
        f'<div class="cell-title" style="margin-top:.45rem">Received: {_e(received)}</div>'
        f'<div class="cell-meta">{_e(received_meta)}</div>'
    )


def _country_sheet_names(country: CountryResult, side: str) -> list[str]:
    attribute = "sent_sheet_names" if side == "sent" else "received_sheet_names"
    names = list(getattr(country, attribute, []) or [])
    if names:
        return names
    indexed: list[tuple[int, str]] = []
    for comparison in country.sheets:
        index = comparison.sent_index if side == "sent" else comparison.received_index
        name = comparison.sent_name if side == "sent" else comparison.received_name
        if index is not None and name is not None:
            indexed.append((index, name))
    return [name for _, name in sorted(indexed)]


def _inventory_list_html(
    names: list[str],
    other_names: list[str],
    side: str,
    reference_counts: Mapping[str, int],
) -> str:
    if not names:
        return '<p class="muted">No sheet names available.</p>'
    other_positions = {name: index for index, name in enumerate(other_names, start=1)}
    items: list[str] = []
    for index, name in enumerate(names, start=1):
        if name not in other_positions:
            state = _badge("Deleted" if side == "sent" else "Added", "error", "!")
        elif other_positions[name] != index:
            state = _badge(f"Position {other_positions[name]}", "info")
        else:
            state = ""
        ref_count = reference_counts.get(name)
        ref_text = (
            f'<div class="cell-meta">#REF! {side.title()} {ref_count:,}</div>'
            if ref_count is not None
            else '<div class="cell-meta">#REF! count unavailable</div>'
        )
        items.append(
            f'<li><span class="inventory-name">{_e(name)}</span>{state}{ref_text}</li>'
        )
    return f'<ol class="inventory-list">{"".join(items)}</ol>'


def _sheet_inventory_html(country: CountryResult) -> str:
    sent_names = _country_sheet_names(country, "sent")
    received_names = _country_sheet_names(country, "received")
    names_match = set(sent_names) == set(received_names)
    order_matches = sent_names == received_names
    sent_ref_counts: dict[str, int] = {}
    received_ref_counts: dict[str, int] = {}
    for comparison in country.sheets:
        if comparison.sent_name is not None and comparison.sent_metrics is not None:
            sent_ref_counts[comparison.sent_name] = int(
                getattr(
                    comparison.sent_metrics,
                    "ref_error_count",
                    getattr(comparison.sent_metrics, "ref_error_cells", 0),
                )
            )
        if comparison.received_name is not None and comparison.received_metrics is not None:
            received_ref_counts[comparison.received_name] = int(
                getattr(
                    comparison.received_metrics,
                    "ref_error_count",
                    getattr(comparison.received_metrics, "ref_error_cells", 0),
                )
            )
    if names_match and order_matches:
        status_label = "Names and order match"
        status_badge = _status_badge("OK")
    elif names_match:
        status_label = "Names match; order differs"
        status_badge = _badge("ORDER DIFFERS", "warn", "!")
    else:
        status_label = "Sheet names differ"
        status_badge = _badge("NAMES DIFFER", "error", "!")
    open_attr = "" if order_matches else " open"
    return f"""
<section id="sheet-inventory-section" class="section" aria-labelledby="sheet-inventory-heading">
  <div class="section__heading"><h2 id="sheet-inventory-heading">Sheet inventory</h2></div>
  <p class="section-intro">Exact, case-sensitive Excel sheet names are listed in workbook order. Counts alone do not determine whether the inventories match.</p>
  <details id="sheet-inventory" class="card sheet-inventory"{open_attr}>
    <summary><span class="sheet-inventory__summary"><span>{_e(status_label)} · {len(sent_names):,} sent → {len(received_names):,} received</span>{status_badge}</span></summary>
    <div class="sheet-inventory__body inventory-grid">
      <div class="inventory-panel"><h3>Sent workbook ({len(sent_names):,})</h3>{_inventory_list_html(sent_names, received_names, 'sent', sent_ref_counts)}</div>
      <div class="inventory-panel"><h3>Received workbook ({len(received_names):,})</h3>{_inventory_list_html(received_names, sent_names, 'received', received_ref_counts)}</div>
    </div>
  </details>
</section>
"""


def _sheet_table_html(country: CountryResult, finding_anchors: Mapping[str, str]) -> str:
    if not country.sheets:
        return (
            '<div class="notice notice--info"><p class="notice__title">No sheet comparison available</p>'
            "<p>The files were not paired, could not be read, or did not produce worksheet evidence.</p></div>"
        )
    rows: list[str] = []
    for index, comparison in enumerate(country.sheets, start=1):
        links = [
            f'<a href="#{_e(finding_anchors[finding_id])}">{_e(finding_id)}</a>'
            for finding_id in comparison.anomaly_ids
            if finding_id in finding_anchors
        ]
        finding_links = (
            '<div class="cell-meta">Findings: ' + ", ".join(links) + "</div>"
            if links
            else ""
        )
        rows.append(
            f'<tr id="sheet-{index}" class="sheet-summary-row">'
            f"<td>{_status_badge(comparison.status)}{finding_links}</td>"
            f"<th scope=\"row\">{_sheet_name_html(comparison)}</th>"
            f"<td>{_metric_html(comparison.sent_metrics)}</td>"
            f"<td>{_metric_html(comparison.received_metrics)}</td>"
            f"<td>{_extent_delta(comparison)}</td>"
            f"<td>{_axis_operations_html(comparison.row_operations, 'ROW', comparison.status)}</td>"
            f"<td>{_axis_operations_html(comparison.column_operations, 'COLUMN', comparison.status)}</td>"
            f'<td class="reference-errors cell-integrity">{_reference_errors_html(comparison)}'
            f'<div class="kpi-evidence">{_cell_integrity_html(comparison)}</div></td>'
            "</tr>"
            '<tr class="sheet-evidence-row"><td colspan="8">'
            f"{_sheet_evidence_html(comparison)}</td></tr>"
        )
    return (
        '<div class="table-wrap" tabindex="0" role="region" aria-label="Sheet comparison matrix">'
        '<table class="sheet-table"><caption>Active extents, #REF! counts, aligned formula/value changes, and inferred operations are shown per sheet; expanded evidence uses the full table width.</caption>'
        "<thead><tr><th scope=\"col\">Status</th><th scope=\"col\">Sheet</th>"
        "<th scope=\"col\">Sent active extent</th><th scope=\"col\">Received active extent</th>"
        "<th scope=\"col\">Extent delta</th><th scope=\"col\">Row operations</th>"
        "<th scope=\"col\">Column operations</th>"
        "<th scope=\"col\">#REF! cells / formula / value changes</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


def _country_kpis(country: CountryResult) -> str:
    metrics = country.metrics
    comparable = country.comparison_state == "PAIRED"
    high_count = sum(finding.severity.upper() == "HIGH" for finding in country.findings)
    medium_count = sum(finding.severity.upper() == "MEDIUM" for finding in country.findings)
    if comparable:
        sheets = f"{metrics.sent_sheet_count:,} → {metrics.received_sheet_count:,}"
        sheet_help = _operation_summary(metrics.sheets_added, metrics.sheets_deleted)
        sheet_help += (
            "; names match exactly"
            if getattr(metrics, "sheet_names_match", True)
            else "; names differ"
        )
        rows = f"+{metrics.rows_added:,} / −{metrics.rows_deleted:,}"
        row_help = f"Net {_signed(metrics.row_net_delta)}"
        columns = f"+{metrics.columns_added:,} / −{metrics.columns_deleted:,}"
        column_help = f"Net {_signed(metrics.column_net_delta)}"
        sent_refs = int(getattr(metrics, "sent_ref_errors", 0))
        received_refs = int(getattr(metrics, "received_ref_errors", 0))
        refs = f"{sent_refs:,} → {received_refs:,}"
        ref_help = f"Delta {_signed(received_refs - sent_refs)}"
        sent_kpis = int(getattr(metrics, "kpi_sent_count", 0))
        received_kpis = int(getattr(metrics, "kpi_received_count", 0))
        kpis = f"{sent_kpis:,} → {received_kpis:,}"
        kpi_help = (
            f"{int(getattr(metrics, 'kpi_missing_count', 0)):,} missing / "
            f"{int(getattr(metrics, 'kpi_unexpected_count', 0)):,} unexpected"
        )
        formula_changes = int(getattr(metrics, "formula_changed_count", 0))
        formula_unresolved = int(getattr(metrics, "formula_unresolved_count", 0))
        formula_help = f"{formula_unresolved:,} skipped/unresolved"
        value_changes = int(getattr(metrics, "value_changed_count", 0))
        value_unresolved = int(getattr(metrics, "value_unresolved_count", 0))
        value_help = f"{value_unresolved:,} skipped/unresolved"
    else:
        sheets = rows = columns = refs = kpis = formula_changes = value_changes = "—"
        formula_unresolved = value_unresolved = 0
        sheet_help = row_help = column_help = ref_help = kpi_help = formula_help = value_help = "Not compared"
    result_tone = (
        "error"
        if country.overall_status.upper() == "ERROR"
        else "warn"
        if country.overall_status.upper() in {"WARNING", "WARN"}
        else "ok"
    )
    affected_tone = "error" if high_count else "warn" if medium_count else ""
    return (
        _kpi("Result", country.overall_status, country.comparison_state.replace("_", " "), result_tone)
        + _kpi("HIGH findings", _number(high_count), "Immediate action", "error" if high_count else "")
        + _kpi("MEDIUM findings", _number(medium_count), "Review recommended", "warn" if medium_count else "")
        + _kpi("Sheets", sheets, sheet_help, "accent")
        + _kpi("Rows + / −", rows, row_help, "error" if metrics.rows_added or metrics.rows_deleted else "")
        + _kpi("Columns + / −", columns, column_help, "error" if metrics.columns_added or metrics.columns_deleted else "")
        + _kpi("#REF! cells", refs, ref_help, "error" if comparable and int(getattr(metrics, "received_ref_errors", 0)) else "")
        + _kpi("KPI identifiers", kpis, kpi_help, "error" if comparable and (int(getattr(metrics, "kpi_missing_count", 0)) or int(getattr(metrics, "kpi_unexpected_count", 0))) else "")
        + _kpi("Formula changes", formula_changes, formula_help, "warn" if comparable and (formula_changes or formula_unresolved) else "")
        + _kpi("Prefilled value changes", value_changes, value_help, "warn" if comparable and (value_changes or value_unresolved) else "")
        + _kpi("Affected sheets", _number(metrics.affected_sheet_count), "With one or more findings", affected_tone if metrics.affected_sheet_count else "")
    )


def _render_country_report(
    run: RunResult,
    country: CountryResult,
    report_filename: str,
    previous_filename: str | None,
    next_filename: str | None,
) -> str:
    anchors_by_object = {
        id(finding): f"finding-{index}" for index, finding in enumerate(country.findings, start=1)
    }
    anchors_by_id: dict[str, str] = {}
    for finding in country.findings:
        anchors_by_id.setdefault(finding.id, anchors_by_object[id(finding)])
    if country.comparison_state == "PAIRED":
        empty_message = "No structural differences were detected by this run."
    else:
        empty_message = "Structural checks were not completed because a comparable workbook pair was unavailable."
    error_notices = _notice("Comparison errors", country.errors, "error")
    warning_notices = _collapsible_notice("Analysis notes", country.warnings, "warn")
    warning_section = (
        '<section id="analysis-notes" class="section" aria-labelledby="analysis-notes-heading">'
        '<div class="section__heading"><h2 id="analysis-notes-heading">Analysis notes</h2></div>'
        f"{warning_notices}</section>"
        if warning_notices
        else ""
    )
    anomaly_filters = _anomaly_filters_html(country.findings)
    anomaly_groups = _finding_groups_html(
        country.findings,
        anchors_by_object,
        empty_message,
    )
    previous_link = (
        f'<a class="button-link" href="{_e(previous_filename)}" rel="prev">Previous country</a>'
        if previous_filename
        else ""
    )
    next_link = (
        f'<a class="button-link" href="{_e(next_filename)}" rel="next">Next country</a>'
        if next_filename
        else ""
    )
    body = f"""
<header class="topbar">
  <div class="topbar__inner">
    <p class="brand">POPS anomaly detection</p>
    <nav class="breadcrumbs" aria-label="Breadcrumb"><a href="../index.html">Global report</a><span aria-hidden="true">/</span><span aria-current="page">{_e(country.display_name)}</span></nav>
  </div>
</header>
<main id="main-content" class="page">
  <nav class="report-nav screen-only" aria-label="Report navigation">
    <a class="button-link" href="../index.html">Back to global report</a>
    <span class="report-nav__spacer"></span>{previous_link}{next_link}
  </nav>
  <div class="hero">
    <div>
      <p class="eyebrow">Country anomaly report</p>
      <h1>{_e(country.display_name)}</h1>
      <p class="lede">{_e(country.sent_file.name if country.sent_file else 'No sent workbook')} → {_e(country.received_file.name if country.received_file else 'No received workbook')}. This report covers workbook structure, KPI identifiers, #REF! errors, aligned formulas, and aligned prefilled values.</p>
    </div>
    <div class="hero__status" aria-label="Country result">
      {_status_badge(country.overall_status)}
      {_status_badge(country.comparison_state)}
      {_severity_badge(country.max_anomaly_severity)}
    </div>
  </div>
  {error_notices}
  <section class="section" aria-labelledby="summary-heading">
    <div class="section__heading"><h2 id="summary-heading">Summary</h2><a href="{_e(Path(report_filename).with_suffix('.json').name)}">Download country JSON</a></div>
    <div class="grid kpi-grid">{_country_kpis(country)}</div>
  </section>
  {_sheet_inventory_html(country)}
  <div id="anomaly-results" aria-label="Anomaly findings">
    {anomaly_filters}
    {anomaly_groups}
  </div>
  <section class="section" aria-labelledby="sheets-heading">
    <div class="section__heading"><h2 id="sheets-heading">Sheet comparison</h2></div>
    <p class="section-intro">Active extents, inferred row and column operations, KPI evidence, #REF! counts, and aligned formula/value change counts are shown for every available sheet comparison.</p>
    {_sheet_table_html(country, anchors_by_id)}
  </section>
  <section class="section" aria-labelledby="sources-heading">
    <div class="section__heading"><h2 id="sources-heading">File evidence</h2></div>
    <div class="grid two-col">{_file_card('Sent workbook', country.sent_file)}{_file_card('Received workbook', country.received_file)}</div>
  </section>
  <section class="section" aria-labelledby="method-heading">
    <div class="section__heading"><h2 id="method-heading">Methodology and limitations</h2></div>
    <div class="card">
      <ul class="method-list">
        <li>Sheet names are compared exactly. A renamed sheet is conservatively represented as a deletion and an addition.</li>
        <li>Active extents are reconstructed from OOXML structural evidence; Excel's declared dimension is displayed as evidence but is not trusted as the sole source.</li>
        <li>Row and column operations are inferred through structural sequence alignment. Recalculated array/data-table outputs and refresh-driven PivotTable/query-table cells and ranges are excluded from those signatures. Confidence and supporting evidence are shown for auditability; an unresolved alignment is itself a HIGH finding.</li>
        <li>Extent deltas are informational. Structural status uses detected operations, so an addition and deletion can be reported even when the final row or column count is unchanged.</li>
        <li>Only the sheet named KPI is evaluated for KPI identifiers. The detector resolves the column headed KPI and compares normalized identifiers, duplicates, and order.</li>
        <li>#REF! totals count unique cells; cached error values and formula-text references are shown separately because one cell can provide both forms of evidence. Only an increase in explicit formula tokens creates a HIGH finding; cache-only changes remain metrics.</li>
        <li>Formula presence and normalized formula logic are compared after logical row and column alignment. Cached results are ignored. Cells inside structurally removed axes, and sheets whose mappings are unresolved, are skipped.</li>
        <li>Meaningful prefilled values from the sent workbook are compared at their logically aligned received cells. Blank, zero, and lone-hyphen placeholders, plus stored array-formula, data-table, PivotTable, query-table, and calculated-column/total outputs, are excluded. Structurally removed or unresolved cells are skipped rather than reported as content changes.</li>
        <li>Formatting intent, charts, and broader business-rule validation remain outside this report's current scope.</li>
      </ul>
    </div>
  </section>
  <section class="section" aria-labelledby="audit-heading">
    <div class="section__heading"><h2 id="audit-heading">Report evidence</h2></div>
    <div class="card"><dl class="audit-dl">
      <dt>Run ID</dt><dd><code>{_e(run.run_id)}</code></dd>
      <dt>Country ID</dt><dd><code>{_e(country.country_id)}</code></dd>
      <dt>Generated</dt><dd><time datetime="{_e(run.generated_at_utc)}">{_e(run.generated_at_utc)}</time></dd>
      <dt>Comparator</dt><dd>{_e(run.comparator_version)}</dd>
      <dt>Report schema</dt><dd>{_e(run.schema_version)}</dd>
    </dl></div>
  </section>
  {warning_section}
  <p class="print-only">Country {_e(country.display_name)} · run {_e(run.run_id)} · generated {_e(run.generated_at_utc)}</p>
</main>
<footer class="footer"><div class="footer__inner">POPS anomaly detector · <a href="../index.html">Global report</a> · Run <code>{_e(run.run_id)}</code></div></footer>
"""
    return _document(
        f"{country.display_name} — POPS anomaly report",
        body,
        _ANOMALY_FILTER_JS + _PRINT_DETAILS_JS,
    )


def render_country_report(run: RunResult, country: CountryResult) -> str:
    """Render one country HTML report without writing it to disk."""

    filename = _safe_filename(country.report_filename, country.country_id or country.display_name)
    return _render_country_report(run, country, filename, None, None)


def write_reports(run: RunResult, output_dir: str | Path) -> Path:
    """Write global and per-country HTML/JSON reports and return the output directory.

    Existing unrelated files in ``output_dir`` are left untouched.  Each individual
    file is replaced atomically after its complete content has been rendered.
    """

    output = Path(output_dir)
    countries_dir = output / "countries"
    output.mkdir(parents=True, exist_ok=True)
    countries_dir.mkdir(parents=True, exist_ok=True)
    filenames = _filename_map(run.countries)

    run_json = json.dumps(
        _run_payload(run, filenames), ensure_ascii=False, indent=2, sort_keys=False
    ) + "\n"
    _atomic_write(output / "report.json", run_json)

    for index, country in enumerate(run.countries):
        filename = filenames[index]
        previous_filename = filenames[index - 1] if index > 0 else None
        next_filename = filenames[index + 1] if index + 1 < len(run.countries) else None
        html = _render_country_report(
            run,
            country,
            filename,
            previous_filename,
            next_filename,
        )
        country_json = json.dumps(
            _country_payload(country, filename),
            ensure_ascii=False,
            indent=2,
            sort_keys=False,
        ) + "\n"
        _atomic_write(countries_dir / filename, html)
        _atomic_write(countries_dir / Path(filename).with_suffix(".json"), country_json)

    _atomic_write(output / "index.html", _render_global_report(run, filenames))
    return output
