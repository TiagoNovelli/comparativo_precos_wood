# Tabela de Preços Wood — Private Label (SoHome)

This is mostly a **data/spreadsheet project**, not a software codebase — pricing tables
for the "Wood" furniture line (Grupo SoHome), two static HTML comparison reports, and
one Python script (`gerar_comparativo.py`) that regenerates those reports from the
spreadsheets. There is no build process, package manager, or test suite.

## Files

- `1 - TABELA DE PRECOS WOOD - PRIVATE LABEL 28.04.xlsx` — main price table for the
  Private Label brand (this folder's namesake). Sheets per furniture category
  (MESA DE JANTAR, MESA DE CENTRO, APARADOR, BUFFET, CÔMODA, RACK, DESK, CAR BAR,
  PENTEADEIRA, etc.), plus CAPA (cover), CALCULO (pricing formulas/coefficients),
  BASE DE DADOS PREÇOS, POLITICA DE PREÇO (payment terms & discount coefficients),
  TIPOS DE TAMPO (tabletop finish types), ESPELHO.
- `1 - TABELA DE PRECOS WOOD - CENTURY 28.04.xlsx` / `... - PV 28.04.xlsx` — equivalent
  price tables for the Century and PV brands (same sheet layout, fewer categories).
- `Novos preços.xlsx` — flat export ("Export" sheet) of new prices: one row per
  model/configuration, columns per finish (LAMINADO, LAQUEADA, MARMORE ESCOVADO, ESP, etc.).
- `Preços antigos - Century e .pv.xlsx` — flat export of old Century/PV prices
  (modelo, modulação, configuração, preço).
- `Comparativo Preços Wood.html` — standalone, self-contained HTML report (inline
  CSS/JS, no external deps besides Google Fonts) comparing old vs. new prices by
  category and finish. Open directly in a browser.
- `Comparativo Preços Wood - Mármore ESP.html` — variant of the comparison report
  that includes "Mármore Especial (ESP)" finishes (excluded from the main report).
- `gerar_comparativo.py` — Python script that regenerates both HTML reports straight
  from the xlsx files (see "Regenerating the comparison reports" below).

## Domain notes (important for interpreting/editing data or reports)

- **Acabamento** = finish/material (e.g., Laminado, Laqueada, Mármore Escovado,
  Vidro). **Tampo** = tabletop. **Modulação/Configuração** = product dimensions/specs.
- Pricing corrections applied in the comparison reports (see the `.note` block at
  the top of each HTML):
  - **Mármore Fornecido** → price without top is the reference, minus 8%.
  - **Vidro Normal** → same price as "sem topo" (column "Tampo: Laca ou Lamina com
    ou sem Vidro Normal").
  - **Mármores Especiais (ESP)** → excluded from the main comparison (see separate
    ESP report).
  - **Mesa Eclipse** → old pricing covered only the upper top; new pricing covers
    upper + lower top. Reports show cheapest/most-expensive config per category.
  - **Sem correspondente** (no match) → products with no equivalent in the old
    tables are listed at the end with new prices only, no comparison.
- Old price sources referenced: Century, PV, Private Label 28/04, and "Preços Antigos".

## Regenerating the comparison reports

`gerar_comparativo.py` reads `Novos preços.xlsx` (new prices) plus the three brand
workbooks (Century / PV / Private Label — old prices) and rewrites both HTML reports.
Run it with `python gerar_comparativo.py` (requires `openpyxl`); it overwrites the two
`Comparativo Preços Wood*.html` files in place.

Matching rule (deterministic, documented at the top of the script): for each new
product, search the old brand sheets of the **same category** for a product whose name
shares its keyword, then pick the (new size variant, old product) pair with the
**smallest combined dimension difference** (C+L+A). No match in-category → the product
goes to "Sem Correspondente". This reproduces the report's structure, corrections
(Mármore Fornecido −8%, Vidro Normal = "sem topo"), and visual style, but **will not
reproduce the previous HTML byte-for-byte** — the original involved manual/curated
choices (e.g., picking specific size variants, an "Eclipse" combinatorial pricing
block, and a legacy `Preços antigos - Century e .pv.xlsx` source) that aren't fully
deterministic. These limitations are spelled out in the script's module docstring.

## Working with these files

- xlsx files are Excel workbooks with multiple sheets per category — open with Excel
  or a library like `openpyxl` (Python) for inspection/edits.
- The HTML reports are static, single-file documents — edit the inline `<style>`/HTML/
  the small `go()` JS helper directly; there's no templating or asset pipeline.
- Filenames and sheet names contain accented Portuguese characters (preços, cômoda,
  cabeceira) — when scripting against these files, use UTF-8-aware tools.
