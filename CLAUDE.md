# Tabela de Preços Wood — Private Label (SoHome)

This is mostly a **data/spreadsheet project**, not a software codebase — pricing tables
for the "Wood" furniture line (Grupo SoHome), one self-contained HTML comparison report,
and one Python script (`gerar_comparativo.py`) that regenerates that report from the
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
- `Preços antigos - Century e .pv.xlsx` — flat export ("Export" sheet) of legacy
  Century/PV prices: one row per model/modulação/configuração, with the **tabletop
  finish type described inside free-text fields** (e.g. "TIPO DE ACABAMENTO DO TAMPO",
  "OPCAO DE TOPO", "TIPO DE ACABAMENTO DA MOLDURA") rather than in fixed columns like
  the three brand workbooks. Used as a **fallback matching source** — see "Regenerating
  the comparison reports" below.
- `Comparativo Preços Wood.html` — standalone, self-contained HTML report (inline
  CSS/JS, no external deps besides Google Fonts) comparing old vs. new prices by
  category and finish, with a client-side **switch to toggle Mármore Especial (ESP)
  finishes in/out of the analysis and percentage calculations** (both views are
  pre-computed server-side and embedded; the switch just shows/hides via CSS). Open
  directly in a browser.
- `gerar_comparativo.py` — Python script that regenerates the HTML report straight
  from the xlsx files (see "Regenerating the comparison reports" below).

## Domain notes (important for interpreting/editing data or reports)

- **Acabamento** = finish/material (e.g., Laminado, Laqueada, Mármore Escovado,
  Vidro). **Tampo** = tabletop. **Modulação/Configuração** = product dimensions/specs.
- Pricing corrections applied in the comparison report (see the `.note` block at
  the top of the HTML):
  - **Mármore Fornecido** → price without top is the reference, minus 8%.
  - **Vidro Normal** → same price as "sem topo" (column "Tampo: Laca ou Lamina com
    ou sem Vidro Normal").
  - **Mármores Especiais (ESP)** → included or excluded from the comparison (and from
    the percentage calculations) via the report's client-side ESP switch — both views
    are pre-computed and embedded, so no regeneration is needed to flip between them.
  - **Mesa Eclipse** → old pricing covered only the upper top; new pricing covers
    upper + lower top. Reports show cheapest/most-expensive config per category.
  - **Sem correspondente** (no match in any source) → listed at the end with only the
    new prices for **one representative configuration** (not every available size/
    finish variant): nearest-to-2.70m for Mesa de Jantar, otherwise the variant whose
    combined dimensions (C+L+A) are closest to the product's average.
  - **Linhas informativas** (no `%`/`R$` delta, shown as "—") → whenever a finish
    exists on only one side of the comparison (e.g. a mirror-top price exists in the
    new table but not the old, or vice versa); the available value is listed for
    reference only.
- Old price sources referenced: Century, PV, Private Label 28/04, and the legacy
  "Preços Antigos — Century e .pv" export.

## Regenerating the comparison report

`gerar_comparativo.py` reads `Novos preços.xlsx` (new prices), the three brand
workbooks (Century / PV / Private Label — old prices), and the legacy
`Preços antigos - Century e .pv.xlsx` export, then rewrites the single consolidated
HTML report. Run it with `python gerar_comparativo.py` (requires `openpyxl`); it
overwrites `Comparativo Preços Wood.html` in place.

**Two-stage matching rule** (deterministic, documented at the top of the script):
1. **Brand workbooks (primary):** for each new product, search the old brand sheets of
   the **same category** for a product whose name shares its keyword, then pick the
   (new size variant, old product) pair with the **smallest combined dimension
   difference** (C+L+A).
2. **Legacy export (fallback only, no category filter):** if no brand-workbook match is
   found, apply the same name-keyword + smallest-dimension-difference rule against the
   legacy "Preços antigos - Century e .pv" products — whose finish types are parsed out
   of free-text config fields (matched generically via any key containing "ACABAMENTO"
   or "OPCAO DE TOPO") and compared **directly by (normalized) finish name** against the
   new table's finish columns, with a small alias table for known spelling differences
   (e.g. "VIDRO EXTRACLEAR 4MM" ↔ "VIDRO EXTRACLEAR"). This is also where new categories
   absent from the brand workbooks (Espelho, Mancebo, Sofá Table, etc.) get matched.
   - **Grouped "NENHUM"**: products whose new-table price is only filled under the
     "NENHUM" finish (i.e. the model doesn't differentiate by finish — e.g. SAMMY)
     have that single price compared against **every** finish found in the matched
     legacy configuration, producing one comparison row per old finish.
3. Still no match → the product goes to "Sem Correspondente" (see above).

This reproduces the report's structure, corrections (Mármore Fornecido −8%, Vidro
Normal = "sem topo", ESP toggle), and visual style, but **will not reproduce the
original two-file HTML byte-for-byte** — the original involved manual/curated choices
(e.g. an "Eclipse" combinatorial pricing block) that aren't fully deterministic. These
limitations are spelled out in the script's module docstring.

## Working with these files

- xlsx files are Excel workbooks with multiple sheets per category — open with Excel
  or a library like `openpyxl` (Python) for inspection/edits.
- The HTML report is a static, single-file document — edit the inline `<style>`/HTML/
  the small `go()` (category nav) and `toggleEsp()` (ESP switch) JS helpers directly;
  there's no templating or asset pipeline.
- Filenames and sheet names contain accented Portuguese characters (preços, cômoda,
  cabeceira) — when scripting against these files, use UTF-8-aware tools.
